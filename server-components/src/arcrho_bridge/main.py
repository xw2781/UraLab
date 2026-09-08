import argparse
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

_MODULE_ROOT = Path(__file__).resolve().parent
_SOURCE_ROOT = _MODULE_ROOT.parent
_PRODUCT_ROOT = _SOURCE_ROOT.parent
# Standalone canonical modules (the durable-job lease) live beside the public
# Python API. A frozen Bridge gets them as hidden imports instead.
_REPO_CANONICAL_ROOT = _PRODUCT_ROOT.parent / "python-api" / "src"
_BUNDLE_ROOT = Path(getattr(sys, "_MEIPASS", _MODULE_ROOT)).resolve()

for _path in (_PRODUCT_ROOT, _SOURCE_ROOT, _REPO_CANONICAL_ROOT, _BUNDLE_ROOT):
    if not _path.exists():
        continue
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

try:
    from src.arcrho_bridge.bridge_utils import (
        RESQ_WINDOW_TITLE,
        heartbeat_payload,
        list_instance_files,
        list_json_files_by_mtime,
        read_json,
        remove_old_instances,
        safe_remove,
        window_is_active,
        write_json,
    )
    from src.arcrho_bridge.resq_import_contract import (
        load_resq_reserving_class_import_contract,
    )
    from src.arcrho_bridge.resq_sync_contract import (
        load_resq_reserving_class_sync_contract,
    )
    from src.arcrho_bridge.resq_client import ResQClient
    from src.utils import get_config_value, get_project_root, normalize_function_name, resolve_app_path
except ModuleNotFoundError:
    from arcrho_bridge.bridge_utils import (
        RESQ_WINDOW_TITLE,
        heartbeat_payload,
        list_instance_files,
        list_json_files_by_mtime,
        read_json,
        remove_old_instances,
        safe_remove,
        window_is_active,
        write_json,
    )
    from arcrho_bridge.resq_import_contract import (
        load_resq_reserving_class_import_contract,
    )
    from arcrho_bridge.resq_sync_contract import (
        load_resq_reserving_class_sync_contract,
    )
    from arcrho_bridge.resq_client import ResQClient
    from utils import get_config_value, get_project_root, normalize_function_name, resolve_app_path


os.environ.setdefault("ARCRHO_ROOT", str(get_project_root()))


RESQ_IMPORT_CONTRACT = load_resq_reserving_class_import_contract()
_RESQ_IMPORT_REQUEST_RELATIVE_DIR = RESQ_IMPORT_CONTRACT["request_relative_dir"]
_RESQ_IMPORT_STATUS_RELATIVE_DIR = RESQ_IMPORT_CONTRACT["status_relative_dir"]
_RESQ_IMPORT_HEARTBEAT_RELATIVE_DIR = RESQ_IMPORT_CONTRACT[
    "worker_heartbeat_relative_dir"
]

BRIDGE_ROLE = "bridge"
WORKER_ROLE = RESQ_IMPORT_CONTRACT["worker_role"]
REQUEST_SUBDIR = _RESQ_IMPORT_REQUEST_RELATIVE_DIR[1]
WORKER_STALE_AFTER_SECONDS = RESQ_IMPORT_CONTRACT[
    "worker_heartbeat_max_age_seconds"
]
REQUEST_POLL_INTERVAL_SECONDS = 1.0
IMPORT_HEARTBEAT_INTERVAL_SECONDS = 1.0
# A running import republishes its own status at least this often, so a
# ``processing`` status older than the threshold has no live owner: its worker
# was terminated without running any exception handler. Startup reconciliation
# closes those out; otherwise the UI waits on an import that will never report.
RESQ_IMPORT_STATUS_STALE_SECONDS = 120.0
RESQ_IMPORT_ORPHANED_STATUS_MESSAGE = (
    "The ArcRho Bridge stopped before this import finished. The live reserving "
    "class was left unchanged; import it again."
)

# A reserving-class import is intentionally isolated from both the legacy RPC
# queue and the data-engine's top-level ``requests`` queue. The latter is
# watched by ArcRho Engine workers, which must never claim a ResQ import.
RESQ_IMPORT_FUNCTION = RESQ_IMPORT_CONTRACT["function"]
RESQ_IMPORT_CONTRACT_VERSION = RESQ_IMPORT_CONTRACT["contract_version"]
_RESQ_IMPORT_REQUIRED_FIELDS = RESQ_IMPORT_CONTRACT["required_request_fields"]
_RESQ_IMPORT_FORBIDDEN_PATH_FIELDS = RESQ_IMPORT_CONTRACT["forbidden_path_fields"]
_RESQ_IMPORT_ALLOWED_EXPORT_MODES = frozenset(
    RESQ_IMPORT_CONTRACT["allowed_export_modes"]
)
_RESQ_IMPORT_STATUS_VALUES = frozenset(RESQ_IMPORT_CONTRACT["status_values"])
_RESQ_IMPORT_REQUEST_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$")
_RESQ_IMPORT_INVALID_PROJECT_NAME_CHARS = frozenset('<>:"/\\|?*\x00')

# The synchronization queue is a sibling of the import queue served by this same
# ResQ-connected worker. Its contract restates none of the worker's identity or
# the queue's status vocabulary; it reads both from the import contract.
RESQ_SYNC_CONTRACT = load_resq_reserving_class_sync_contract()
_RESQ_SYNC_REQUEST_RELATIVE_DIR = RESQ_SYNC_CONTRACT["request_relative_dir"]
_RESQ_SYNC_STATUS_RELATIVE_DIR = RESQ_SYNC_CONTRACT["status_relative_dir"]
RESQ_SYNC_FUNCTION = RESQ_SYNC_CONTRACT["function"]
RESQ_SYNC_CONTRACT_VERSION = RESQ_SYNC_CONTRACT["contract_version"]
_RESQ_SYNC_REQUIRED_FIELDS = RESQ_SYNC_CONTRACT["required_request_fields"]
_RESQ_SYNC_ALLOWED_PHASES = frozenset(RESQ_SYNC_CONTRACT["allowed_phases"])
_RESQ_SYNC_SELECTION_FIELD = RESQ_SYNC_CONTRACT["selection_field"]
RESQ_SYNC_ORPHANED_STATUS_MESSAGE = (
    "The ArcRho Bridge stopped before this synchronization finished. Compare the "
    "reserving class again to see what was applied."
)


def normalize_method_name(method_name):
    return re.sub(r"\s+", " ", str(method_name or "")).strip()


def make_instance_id(role):
    device_name = os.environ.get("COMPUTERNAME", "UNKNOWN")
    ts = datetime.now().strftime("%y%m%d-%H%M%S-%f")[:-3]
    return f"{role}@{device_name}@{os.getlogin()}@{ts}"


def instance_path(role, instance_id):
    return resolve_app_path(role, "instances", f"{instance_id}.json")


def request_dir():
    path = get_project_root().joinpath(*_RESQ_IMPORT_REQUEST_RELATIVE_DIR[:2])
    path.mkdir(parents=True, exist_ok=True)
    return path


def resq_import_queue_dir(server_root=None):
    """Return the logical shared-server queue root for ResQ RC imports.

    Callers exchange only logical project and reserving-class identifiers. In
    particular, no producer-local mapped-drive path is accepted in a request;
    each machine resolves this directory from its own ArcRho Server root.
    """

    root = Path(server_root) if server_root is not None else get_project_root()
    return root.joinpath(*_RESQ_IMPORT_REQUEST_RELATIVE_DIR[:-1])


def resq_import_request_dir(server_root=None):
    root = Path(server_root) if server_root is not None else get_project_root()
    path = root.joinpath(*_RESQ_IMPORT_REQUEST_RELATIVE_DIR)
    path.mkdir(parents=True, exist_ok=True)
    return path


def resq_import_status_dir(server_root=None):
    root = Path(server_root) if server_root is not None else get_project_root()
    path = root.joinpath(*_RESQ_IMPORT_STATUS_RELATIVE_DIR)
    path.mkdir(parents=True, exist_ok=True)
    return path


def resq_import_status_path(request_id, server_root=None):
    """Return the deterministic status path for one accepted import request."""

    normalized_id = _validate_resq_import_request_id(request_id)
    return resq_import_status_dir(server_root) / f"{normalized_id}.json"


def resq_sync_request_dir(server_root=None):
    root = Path(server_root) if server_root is not None else get_project_root()
    path = root.joinpath(*_RESQ_SYNC_REQUEST_RELATIVE_DIR)
    path.mkdir(parents=True, exist_ok=True)
    return path


def resq_sync_status_dir(server_root=None):
    root = Path(server_root) if server_root is not None else get_project_root()
    path = root.joinpath(*_RESQ_SYNC_STATUS_RELATIVE_DIR)
    path.mkdir(parents=True, exist_ok=True)
    return path


def resq_sync_status_path(request_id, server_root=None):
    """Return the deterministic status path for one accepted sync request."""

    normalized_id = _validate_resq_import_request_id(request_id)
    return resq_sync_status_dir(server_root) / f"{normalized_id}.json"


def worker_instance_folder():
    path = resolve_app_path(WORKER_ROLE, "instances")
    path.mkdir(parents=True, exist_ok=True)
    return path


def _resq_gui_is_running(value):
    """Accept only explicit true values from a worker heartbeat."""

    if isinstance(value, bool):
        return value
    return str(value or "").strip().casefold() in {"1", "true", "yes"}


def discover_fresh_bridge_worker_heartbeats(
    server_root=None,
    *,
    max_age_seconds=WORKER_STALE_AFTER_SECONDS,
    now=None,
    user=None,
):
    """Return fresh bridge-worker heartbeats without mutating the share.

    Modification time is deliberate: it works on mapped/UNC shares where
    filesystem-watch events can be delayed or dropped, and it does not depend
    on the submitting PC's local drive-letter alias.
    """

    if max_age_seconds < 0:
        raise ValueError("max_age_seconds must be non-negative.")

    root = Path(server_root) if server_root is not None else get_project_root()
    folder = root.joinpath(*_RESQ_IMPORT_HEARTBEAT_RELATIVE_DIR)
    observed_at = time.time() if now is None else float(now)
    normalized_user = str(user).casefold() if user else None
    fresh = []
    for path in list_json_files_by_mtime(folder):
        try:
            age_seconds = observed_at - path.stat().st_mtime
            if age_seconds < -max_age_seconds or age_seconds > max_age_seconds:
                continue
            payload = read_json(path)
        except OSError:
            # A heartbeat may disappear while its supervisor cleans it up.
            continue
        except Exception:
            # An incomplete or malformed heartbeat is not evidence of a live,
            # ResQ-connected worker.
            continue
        if (
            isinstance(payload, dict)
            and payload.get("Role") == WORKER_ROLE
            and _resq_gui_is_running(payload.get("ResQGuiRunning"))
            and (
                normalized_user is None
                or str(payload.get("User") or "").casefold() == normalized_user
            )
        ):
            fresh.append(path)
    return tuple(sorted(fresh, key=lambda item: item.name.casefold()))


def live_worker_count(user=None):
    remove_old_instances(worker_instance_folder(), WORKER_STALE_AFTER_SECONDS)
    return len(discover_fresh_bridge_worker_heartbeats(user=user))


def _instance_file_user(path):
    """Extract the login from ``<role>@<machine>@<user>@<timestamp>`` names."""

    parts = Path(path).stem.split("@")
    return parts[-2] if len(parts) >= 3 else ""


def remove_worker_heartbeats(user=None):
    """Remove worker heartbeats, optionally only those owned by one user.

    Bridges run one per user session on a shared PC, so a supervisor must not
    delete another user's worker heartbeat; that would stop their live worker.
    """

    normalized_user = str(user).casefold() if user else None
    for path in list_instance_files(worker_instance_folder()):
        if (
            normalized_user is not None
            and _instance_file_user(path).casefold() != normalized_user
        ):
            continue
        safe_remove(path)


def worker_command():
    if getattr(sys, "frozen", False):
        return [sys.executable, "--worker"]
    return [sys.executable, str(Path(__file__).resolve()), "--worker"]


def start_worker():
    return subprocess.Popen(worker_command(), close_fds=True)


def stop_worker(process, timeout=2.0):
    if process is None:
        return None
    if process.poll() is not None:
        return None
    process.terminate()
    try:
        process.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=timeout)
    return None


BRIDGE_STALE_AFTER_SECONDS = 60


def same_user_bridge_is_running(user):
    """Return whether a live Bridge heartbeat already exists for this user."""

    folder = resolve_app_path(BRIDGE_ROLE, "instances")
    normalized_user = str(user).casefold()
    now = time.time()
    for path in list_instance_files(folder):
        try:
            if now - path.stat().st_mtime > BRIDGE_STALE_AFTER_SECONDS:
                continue
            payload = read_json(path, retries=3)
        except Exception:
            continue
        if (
            isinstance(payload, dict)
            and payload.get("Role") == BRIDGE_ROLE
            and str(payload.get("User") or "").casefold() == normalized_user
        ):
            return True
    return False


def run_bridge_supervisor():
    # One bridge per user session on a shared PC: each ResQ user contributes
    # the bridge tied to their own ResQ GUI and license. Only a duplicate for
    # the same user exits here; other users' bridges are never touched.
    current_user = os.getlogin()
    if same_user_bridge_is_running(current_user):
        print(f"An ArcRho Bridge is already running for {current_user}; exiting.")
        return

    bridge_id = make_instance_id(BRIDGE_ROLE)
    id_path = instance_path(BRIDGE_ROLE, bridge_id)
    id_path.parent.mkdir(parents=True, exist_ok=True)
    worker_process = None

    print("Bridge ID: " + bridge_id + "\n")
    write_json(id_path, heartbeat_payload(bridge_id, BRIDGE_ROLE, Created=datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    try:
        while True:
            if not id_path.exists() or get_config_value("apps.bridge.kill_all", False):
                remove_worker_heartbeats(current_user)
                worker_process = stop_worker(worker_process)
                safe_remove(id_path)
                break

            gui_running = window_is_active(RESQ_WINDOW_TITLE)
            write_json(
                id_path,
                heartbeat_payload(
                    bridge_id,
                    BRIDGE_ROLE,
                    ResQGuiRunning=gui_running,
                    WorkerPid=worker_process.pid if worker_process and worker_process.poll() is None else None,
                ),
            )

            if worker_process and worker_process.poll() is not None:
                worker_process = None

            if get_config_value("apps.bridge_worker.kill_all", False):
                remove_worker_heartbeats(current_user)
                worker_process = stop_worker(worker_process, timeout=0.5)
                time.sleep(2)
                continue

            if not gui_running:
                remove_worker_heartbeats(current_user)
                worker_process = stop_worker(worker_process, timeout=0.5)
            elif (
                live_worker_count(current_user) < int(get_config_value("apps.bridge.max_workers", 1))
                and worker_process is None
            ):
                worker_process = start_worker()

            time.sleep(2)
    except KeyboardInterrupt:
        worker_process = stop_worker(worker_process)
    finally:
        safe_remove(id_path)


def _validate_resq_import_request_id(request_id):
    """Validate the token shared by a request and its deterministic status."""

    normalized_id = str(request_id or "").strip()
    if not _RESQ_IMPORT_REQUEST_ID_PATTERN.fullmatch(normalized_id):
        raise ValueError(
            "RequestId must contain 1-128 letters, numbers, underscores, or hyphens."
        )
    return normalized_id


def _validate_resq_import_project_name(project_name):
    """Return a one-segment logical ArcRho project identity."""

    if not isinstance(project_name, str):
        raise ValueError("ProjectName must be a string.")
    normalized_name = project_name.strip()
    if not normalized_name:
        raise ValueError("ProjectName is required.")
    if normalized_name in {".", ".."} or any(
        character in normalized_name
        for character in _RESQ_IMPORT_INVALID_PROJECT_NAME_CHARS
    ):
        raise ValueError("ProjectName must be one logical path segment.")
    return normalized_name


def _validate_resq_import_rc_path(rc_path):
    """Return a relative Windows ArcRho reserving-class identity."""

    if not isinstance(rc_path, str):
        raise ValueError("Path must be a string.")
    normalized_path = rc_path.strip().replace("/", "\\")
    if not normalized_path:
        raise ValueError("Path is required.")
    segments = [part.strip() for part in normalized_path.split("\\")]
    if (
        normalized_path.startswith("\\")
        or ":" in normalized_path
        or "\x00" in normalized_path
        or any(part in {"", ".", ".."} for part in segments)
    ):
        raise ValueError(
            "Path must be a relative Windows ArcRho reserving-class path without '..'."
        )
    return normalized_path


def _json_safe_status_value(value):
    """Convert a client callback/result to an atomic JSON status value."""

    try:
        return json.loads(json.dumps(value, default=str))
    except (TypeError, ValueError):
        return str(value)


def _write_resq_import_status(request, status, *, message="", progress=None, result=None):
    """Atomically publish the deterministic import status for ``request``."""

    return _write_queue_status(
        request,
        status,
        publish=_publish_resq_import_status,
        label="import",
        message=message,
        progress=progress,
        result=result,
    )


def _write_resq_sync_status(request, status, *, message="", progress=None, result=None):
    """Atomically publish the deterministic sync status for ``request``."""

    return _write_queue_status(
        request,
        status,
        publish=_publish_resq_sync_status,
        label="sync",
        message=message,
        progress=progress,
        result=result,
    )


def _write_queue_status(request, status, *, publish, label, message="", progress=None, result=None):
    try:
        request_id = _validate_resq_import_request_id(request.get("RequestId"))
    except Exception as exc:
        print(f"(error: could not resolve ResQ {label} status path: {exc})")
        return False
    return publish(
        request_id,
        status,
        message=message,
        progress=progress,
        result=result,
    )


def _publish_resq_import_status(
    request_id,
    status,
    *,
    message="",
    progress=None,
    result=None,
    server_root=None,
):
    """Atomically publish one status document for an accepted import request id."""

    return _publish_queue_status(
        request_id,
        status,
        status_path_of=resq_import_status_path,
        contract_version=RESQ_IMPORT_CONTRACT_VERSION,
        label="import",
        message=message,
        progress=progress,
        result=result,
        server_root=server_root,
    )


def _publish_resq_sync_status(
    request_id,
    status,
    *,
    message="",
    progress=None,
    result=None,
    server_root=None,
):
    """Atomically publish one status document for an accepted sync request id."""

    return _publish_queue_status(
        request_id,
        status,
        status_path_of=resq_sync_status_path,
        contract_version=RESQ_SYNC_CONTRACT_VERSION,
        label="sync",
        message=message,
        progress=progress,
        result=result,
        server_root=server_root,
    )


def _publish_queue_status(
    request_id,
    status,
    *,
    status_path_of,
    contract_version,
    label,
    message="",
    progress=None,
    result=None,
    server_root=None,
):
    """Atomically publish one status document for an accepted request id.

    Both ResQ queues report through the same document shape, so a client that
    can poll an import can poll a synchronization with no second reader.
    """

    if status not in _RESQ_IMPORT_STATUS_VALUES:
        raise ValueError(f"Invalid ResQ {label} status: {status!r}")

    try:
        status_path = status_path_of(request_id, server_root)
    except Exception as exc:
        print(f"(error: could not resolve ResQ {label} status path: {exc})")
        return False

    payload = {
        "contract_version": contract_version,
        "status": status,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "request_id": request_id,
    }
    if message:
        payload["message"] = str(message)
    if progress is not None:
        normalized_progress = _json_safe_status_value(progress)
        payload["progress"] = (
            normalized_progress
            if isinstance(normalized_progress, dict)
            else {"message": str(normalized_progress)}
        )
    if result is not None:
        payload["result"] = _json_safe_status_value(result)

    try:
        if write_json(status_path, payload):
            return True
        print(f"(error: could not write ResQ {label} status to {status_path})")
    except Exception as exc:
        print(f"(error: could not write ResQ {label} status to {status_path}: {exc})")
    return False


def _touch_resq_import_status(status_path):
    """Renew a running import's status mtime without rewriting its payload.

    Startup reconciliation reads that mtime to tell a live import from one
    whose worker was terminated. Progress events alone are too sparse: a single
    slow dataset can leave a healthy import silent for minutes.
    """

    try:
        os.utime(status_path, None)
    except OSError:
        pass


def reconcile_orphaned_resq_import_statuses(
    server_root=None,
    *,
    max_age_seconds=RESQ_IMPORT_STATUS_STALE_SECONDS,
    now=None,
):
    """Close out ``processing`` statuses no live worker is renewing.

    Every ArcRho user's Bridge sees the same shared status folder, so freshness
    is the only safe ownership signal here: another machine's running import
    renews its status well inside the threshold and must never be closed out.
    Returns the request ids this call reported as failed.
    """

    return _reconcile_orphaned_statuses(
        server_root,
        status_dir_of=resq_import_status_dir,
        publish=_publish_resq_import_status,
        message=RESQ_IMPORT_ORPHANED_STATUS_MESSAGE,
        label="import",
        on_reconciled=_discard_abandoned_import_job,
        max_age_seconds=max_age_seconds,
        now=now,
    )


def reconcile_orphaned_resq_sync_statuses(
    server_root=None,
    *,
    max_age_seconds=RESQ_IMPORT_STATUS_STALE_SECONDS,
    now=None,
):
    """Close out ``processing`` synchronizations no live worker is renewing.

    A synchronization owns no staging folder, so nothing has to be reclaimed
    here: the apply phase's own lease expires on its own, and every write it
    completed before the worker died is already reported by its results.
    """

    return _reconcile_orphaned_statuses(
        server_root,
        status_dir_of=resq_sync_status_dir,
        publish=_publish_resq_sync_status,
        message=RESQ_SYNC_ORPHANED_STATUS_MESSAGE,
        label="synchronization",
        on_reconciled=None,
        max_age_seconds=max_age_seconds,
        now=now,
    )


def _reconcile_orphaned_statuses(
    server_root,
    *,
    status_dir_of,
    publish,
    message,
    label,
    on_reconciled=None,
    max_age_seconds=RESQ_IMPORT_STATUS_STALE_SECONDS,
    now=None,
):
    if max_age_seconds < 0:
        raise ValueError("max_age_seconds must be non-negative.")

    try:
        folder = status_dir_of(server_root)
    except Exception as exc:
        print(f"(error: could not open the ResQ {label} status folder: {exc})")
        return ()

    observed_at = time.time() if now is None else float(now)
    reconciled = []
    for path in list_json_files_by_mtime(folder):
        try:
            if observed_at - path.stat().st_mtime <= max_age_seconds:
                continue
            payload = read_json(path)
        except OSError:
            continue
        except Exception:
            # An unreadable status is not evidence of abandoned work.
            continue
        if not isinstance(payload, dict) or payload.get("status") != "processing":
            continue
        request_id = payload.get("request_id")
        try:
            request_id = _validate_resq_import_request_id(request_id)
        except Exception:
            continue
        if publish(
            request_id,
            "error",
            message=message,
            server_root=server_root,
        ):
            reconciled.append(request_id)
            if on_reconciled is not None:
                on_reconciled(request_id, server_root)
    if reconciled:
        print(
            f"Closed {len(reconciled)} interrupted ResQ {label}(s): "
            + ", ".join(reconciled)
        )
    return tuple(reconciled)


def _discard_abandoned_import_job(request_id, server_root=None):
    """Reclaim the staging folder of an import that was just declared dead.

    The staging layout belongs to the import runner, so this delegates rather
    than rebuilding those paths here.
    """

    try:
        from arcrho_bridge.resq_import_runner import discard_abandoned_import_job

        discard_abandoned_import_job(request_id, server_root)
    except Exception as exc:
        # Reporting the interrupted import matters more than reclaiming disk.
        print(f"(warning: could not remove staged import [{request_id}]: {exc})")


class BridgeRequestHandler(FileSystemEventHandler):
    def __init__(
        self,
        client,
        *,
        worker_heartbeat=None,
        heartbeat_interval_sec=IMPORT_HEARTBEAT_INTERVAL_SECONDS,
    ):
        self.client = client
        # Watchdog callbacks run on a separate thread, while ResQ COM belongs
        # to the worker thread. Events therefore only request a main-thread
        # scan; they must never invoke ``process_file`` directly.
        self._scan_requested = threading.Event()
        self._process_lock = threading.Lock()
        self._worker_heartbeat = worker_heartbeat
        self._heartbeat_interval_sec = float(heartbeat_interval_sec)

    def on_moved(self, event):
        if event.is_directory:
            return
        self._request_scan(event.dest_path)

    def on_created(self, event):
        if event.is_directory:
            return
        self._request_scan(event.src_path)

    def consume_scan_request(self):
        """Return whether a watchdog event requested a worker-thread scan."""

        if not self._scan_requested.is_set():
            return False
        self._scan_requested.clear()
        return True

    def wait_for_scan_request(self, timeout):
        """Idle until a watchdog event asks for a scan, or *timeout* elapses.

        The worker loop used to sleep a flat second between scans, so a request
        that arrived just after a scan waited out the rest of that second
        before anyone looked -- half a second on average added to every DFM and
        Result Selection sync a user is watching. Waiting on the event instead
        keeps the same ceiling for the periodic rescan a share needs while
        answering a local event immediately. The flag is left set for
        ``consume_scan_request`` to clear.
        """

        self._scan_requested.wait(timeout)

    def process_pending(self, folder):
        """Claim pending requests in deterministic mtime order.

        The worker thread periodically calls this even when a mapped/UNC
        filesystem does not reliably deliver a watchdog event.
        """

        for path in list_json_files_by_mtime(folder):
            if not self.process_file(path):
                break

    def _request_scan(self, path):
        if str(path).lower().endswith(".json"):
            self._scan_requested.set()

    def process_file(self, path):
        if not self._process_lock.acquire(blocking=False):
            return False
        try:
            return self._process_claimed_file(path)
        finally:
            self._process_lock.release()

    def _process_claimed_file(self, path):
        try:
            request = read_json(path)
        except Exception:
            return True

        if not isinstance(request, dict):
            # A valid JSON value that is not an object cannot have a stable
            # RequestId/status path. Claim it so a malformed queue item cannot
            # block every periodic scan forever.
            safe_remove(path)
            return True

        # Every bridge worker sees the same request directory. Claim first,
        # before validation or output, so exactly one worker processes it.
        if not safe_remove(path):
            return True

        function_name = normalize_function_name(request.get("Function", ""))
        if function_name == RESQ_IMPORT_FUNCTION:
            self._process_resq_import_request(request)
            return True
        if function_name == RESQ_SYNC_FUNCTION:
            self._process_resq_sync_request(request)
            return True

        try:
            if function_name == "DFM":
                request["MethodName"] = normalize_method_name(request.get("MethodName", ""))
                self._validate_request(request)
                self.client.write_dfm_payload(request)
            elif function_name == "SyncDFM":
                request["MethodName"] = normalize_method_name(request.get("MethodName", ""))
                self._validate_request(request)
                self._validate_sync_dfm_request(request)
                self.client.write_sync_dfm_payload(request)
            else:
                self.client.write_error(request, f"Invalid function name: {request.get('Function', '')}")
        except Exception as exc:
            self.client.write_error(request, exc)
        return True

    def _process_resq_import_request(self, request):
        # RequestId is the minimum needed to report rejection. All other
        # protocol validation happens after the processing marker, matching the
        # data-engine's request contract and making claim state observable.
        try:
            _validate_resq_import_request_id(request.get("RequestId"))
        except Exception:
            return

        if not _write_resq_import_status(request, "processing"):
            return

        try:
            self._validate_resq_import_request(request)
        except Exception as exc:
            _write_resq_import_status(request, "error", message=exc)
            return

        def publish_progress(progress):
            _write_resq_import_status(request, "processing", progress=progress)

        heartbeat_stop, heartbeat_thread = self._start_import_heartbeat(request)
        try:
            try:
                result = self.client.write_resq_reserving_class_import(
                    request,
                    progress_callback=publish_progress,
                )
            except Exception as exc:
                status_result = getattr(exc, "status_result", None)
                _write_resq_import_status(
                    request,
                    "error",
                    message=exc,
                    result=status_result if isinstance(status_result, dict) else None,
                )
                return
        finally:
            if heartbeat_stop is not None:
                heartbeat_stop.set()
            if heartbeat_thread is not None:
                heartbeat_thread.join(timeout=IMPORT_HEARTBEAT_INTERVAL_SECONDS)

        _write_resq_import_status(request, "success", result=result)

    def _process_resq_sync_request(self, request):
        """Run one queued ArcRho/ResQ synchronization phase for a client macro.

        The claim, status, and heartbeat protocol is the import queue's, so a
        macro that already knows how to wait for a Bridge import needs no
        second reader to wait for a synchronization.
        """

        try:
            _validate_resq_import_request_id(request.get("RequestId"))
        except Exception:
            return

        if not _write_resq_sync_status(request, "processing"):
            return

        try:
            self._validate_resq_sync_request(request)
        except Exception as exc:
            _write_resq_sync_status(request, "error", message=exc)
            return

        def publish_progress(progress):
            _write_resq_sync_status(request, "processing", progress=progress)

        heartbeat_stop, heartbeat_thread = self._start_import_heartbeat(
            request,
            status_path_of=resq_sync_status_path,
        )
        try:
            try:
                result = self.client.write_resq_reserving_class_sync(
                    request,
                    progress_callback=publish_progress,
                )
            except Exception as exc:
                status_result = getattr(exc, "status_result", None)
                _write_resq_sync_status(
                    request,
                    "error",
                    message=exc,
                    result=status_result if isinstance(status_result, dict) else None,
                )
                return
        finally:
            if heartbeat_stop is not None:
                heartbeat_stop.set()
            if heartbeat_thread is not None:
                heartbeat_thread.join(timeout=IMPORT_HEARTBEAT_INTERVAL_SECONDS)

        _write_resq_sync_status(request, "success", result=result)

    def _start_import_heartbeat(self, request, *, status_path_of=None):
        """Keep the worker and its queued status alive during ResQ COM work."""

        resolve = status_path_of or resq_import_status_path
        try:
            status_path = resolve(request.get("RequestId"))
        except Exception:
            status_path = None
        if not callable(self._worker_heartbeat) and status_path is None:
            return None, None
        beat = self._import_heartbeat_writer(status_path)
        beat()
        stop = threading.Event()

        def keepalive():
            while not stop.wait(self._heartbeat_interval_sec):
                beat()

        thread = threading.Thread(
            target=keepalive,
            name="arcrho-bridge-import-heartbeat",
            daemon=True,
        )
        thread.start()
        return stop, thread

    def _import_heartbeat_writer(self, status_path):
        """Return the one callable that renews both liveness signals."""

        def beat():
            if callable(self._worker_heartbeat):
                try:
                    self._worker_heartbeat()
                except Exception:
                    # A heartbeat write is advisory. The import's own status
                    # writer remains responsible for reporting a real failure.
                    pass
            if status_path is not None:
                _touch_resq_import_status(status_path)

        return beat

    def _validate_resq_import_request(self, request):
        if str(request.get("Function") or "").strip() != RESQ_IMPORT_FUNCTION:
            raise ValueError(f"Function must be {RESQ_IMPORT_FUNCTION}.")

        version = request.get("ContractVersion")
        if isinstance(version, bool) or not isinstance(version, int):
            raise ValueError(
                f"ContractVersion must be the integer {RESQ_IMPORT_CONTRACT_VERSION}."
            )
        if version != RESQ_IMPORT_CONTRACT_VERSION:
            raise ValueError(
                f"Unsupported ContractVersion {version!r}; expected "
                f"{RESQ_IMPORT_CONTRACT_VERSION}."
            )

        _validate_resq_import_request_id(request.get("RequestId"))
        missing = []
        for key in _RESQ_IMPORT_REQUIRED_FIELDS:
            if key in {"Function", "ContractVersion", "RequestId"}:
                continue
            value = request.get(key)
            if not isinstance(value, str) or not value.strip():
                missing.append(key)
            else:
                request[key] = value.strip()
        if missing:
            raise ValueError("Missing request field(s): " + ", ".join(missing))

        request["ProjectName"] = _validate_resq_import_project_name(
            request["ProjectName"]
        )
        request["Path"] = _validate_resq_import_rc_path(request["Path"])
        request["ExportMode"] = request["ExportMode"].casefold()
        if request["ExportMode"] not in _RESQ_IMPORT_ALLOWED_EXPORT_MODES:
            raise ValueError(
                "ExportMode must be one of: "
                + ", ".join(sorted(_RESQ_IMPORT_ALLOWED_EXPORT_MODES))
                + "."
            )

        # A status path is derived from RequestId; accepting a producer-supplied
        # mapped-drive path would make cross-PC imports alias-dependent and
        # reopen an arbitrary-write path.
        supplied_paths = [
            key for key in _RESQ_IMPORT_FORBIDDEN_PATH_FIELDS if request.get(key)
        ]
        if supplied_paths:
            raise ValueError(
                "ResQ import request must not supply path field(s): "
                + ", ".join(supplied_paths)
            )

    def _validate_resq_sync_request(self, request):
        if str(request.get("Function") or "").strip() != RESQ_SYNC_FUNCTION:
            raise ValueError(f"Function must be {RESQ_SYNC_FUNCTION}.")

        version = request.get("ContractVersion")
        if isinstance(version, bool) or not isinstance(version, int):
            raise ValueError(
                f"ContractVersion must be the integer {RESQ_SYNC_CONTRACT_VERSION}."
            )
        if version != RESQ_SYNC_CONTRACT_VERSION:
            raise ValueError(
                f"Unsupported ContractVersion {version!r}; expected "
                f"{RESQ_SYNC_CONTRACT_VERSION}."
            )

        _validate_resq_import_request_id(request.get("RequestId"))
        missing = []
        for key in _RESQ_SYNC_REQUIRED_FIELDS:
            if key in {"Function", "ContractVersion", "RequestId"}:
                continue
            value = request.get(key)
            if not isinstance(value, str) or not value.strip():
                missing.append(key)
            else:
                request[key] = value.strip()
        if missing:
            raise ValueError("Missing request field(s): " + ", ".join(missing))

        request["ProjectName"] = _validate_resq_import_project_name(
            request["ProjectName"]
        )
        request["Path"] = _validate_resq_import_rc_path(request["Path"])
        request["Phase"] = request["Phase"].casefold()
        if request["Phase"] not in _RESQ_SYNC_ALLOWED_PHASES:
            raise ValueError(
                "Phase must be one of: "
                + ", ".join(sorted(_RESQ_SYNC_ALLOWED_PHASES))
                + "."
            )
        # Only the apply phase carries a selection, and it must carry one: a
        # writing request with no reviewed rows is a client bug, not a no-op.
        selection = request.get(_RESQ_SYNC_SELECTION_FIELD)
        if request["Phase"] == "apply":
            if not isinstance(selection, list) or not selection:
                raise ValueError(
                    f"{_RESQ_SYNC_SELECTION_FIELD} must list the accepted review rows."
                )
        elif selection is not None:
            raise ValueError(
                f"A {request['Phase']} request must not supply {_RESQ_SYNC_SELECTION_FIELD}."
            )

        supplied_paths = [
            key for key in _RESQ_IMPORT_FORBIDDEN_PATH_FIELDS if request.get(key)
        ]
        if supplied_paths:
            raise ValueError(
                "ResQ synchronization request must not supply path field(s): "
                + ", ".join(supplied_paths)
            )

    def _validate_request(self, request):
        missing = [
            key
            for key in (
                "Function",
                "ProjectName",
                "Path",
                "MethodName",
                "DataPath",
                "UserName",
            )
            if not request.get(key)
        ]
        if missing:
            raise ValueError("Missing request field(s): " + ", ".join(missing))

    def _validate_sync_dfm_request(self, request):
        self._validate_sync_method_request(request, "SyncDFM")

    def _validate_sync_method_request(self, request, function_name):
        missing = [
            key
            for key in (
                "MethodJsonPath",
                "RPCServerWriteConfirmed",
            )
            if not request.get(key)
        ]
        if missing:
            raise ValueError(f"Missing {function_name} request field(s): " + ", ".join(missing))
        if str(request.get("RPCServerWriteConfirmed", "")).strip().lower() not in {"1", "true", "yes"}:
            raise ValueError(f"{function_name} requires explicit RPC server write confirmation.")


def run_bridge_worker():
    if not window_is_active(RESQ_WINDOW_TITLE):
        return

    worker_id = make_instance_id(WORKER_ROLE)
    id_path = instance_path(WORKER_ROLE, worker_id)
    id_path.parent.mkdir(parents=True, exist_ok=True)

    print("Bridge Worker ID: " + worker_id + "\n")
    client = ResQClient()

    def publish_worker_heartbeat(*, created=False):
        payload = heartbeat_payload(worker_id, WORKER_ROLE, ResQGuiRunning=True)
        if created:
            payload["Created"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        write_json(id_path, payload)

    publish_worker_heartbeat(created=True)
    handler = BridgeRequestHandler(client, worker_heartbeat=publish_worker_heartbeat)
    observer = Observer()
    legacy_request_folder = request_dir()
    import_request_folder = resq_import_request_dir()
    sync_request_folder = resq_sync_request_dir()
    observer.schedule(handler, str(legacy_request_folder), recursive=False)
    observer.schedule(handler, str(import_request_folder), recursive=False)
    observer.schedule(handler, str(sync_request_folder), recursive=False)
    observer.start()
    # A worker that was terminated mid-import left its status claiming
    # ``processing`` forever. Close those out before claiming new work, so the
    # UI stops waiting on an import that no process is running.
    reconcile_orphaned_resq_import_statuses()
    reconcile_orphaned_resq_sync_statuses()
    handler.process_pending(legacy_request_folder)
    handler.process_pending(import_request_folder)
    handler.process_pending(sync_request_folder)
    last_request_scan = time.monotonic()

    try:
        while True:
            if not id_path.exists():
                observer.stop()
                break
            if get_config_value("apps.bridge_worker.kill_all", False):
                safe_remove(id_path)
                observer.stop()
                break
            if not window_is_active(RESQ_WINDOW_TITLE):
                safe_remove(id_path)
                observer.stop()
                break
            client.disconnect_if_idle()
            publish_worker_heartbeat()
            # Watchdog events are opportunistic on a mapped/UNC share. Poll the
            # request folders as well; atomic claim still guarantees that only
            # one bridge worker handles any request.
            if (
                handler.consume_scan_request()
                or time.monotonic() - last_request_scan >= REQUEST_POLL_INTERVAL_SECONDS
            ):
                handler.process_pending(legacy_request_folder)
                handler.process_pending(import_request_folder)
                handler.process_pending(sync_request_folder)
                last_request_scan = time.monotonic()
            handler.wait_for_scan_request(1)
    except KeyboardInterrupt:
        observer.stop()
    finally:
        client.close()
        observer.join()
        safe_remove(id_path)


def main():
    parser = argparse.ArgumentParser(description="Run ArcRho Bridge.")
    parser.add_argument("--worker", action="store_true", help="Run as the ResQ-connected bridge worker.")
    args = parser.parse_args()

    if args.worker:
        run_bridge_worker()
    else:
        run_bridge_supervisor()


if __name__ == "__main__":
    main()
