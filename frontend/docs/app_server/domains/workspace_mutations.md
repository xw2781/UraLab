# App Server Domain: Workspace Mutations (Server-Hosted Transport)

## Purpose
<!-- MANUAL:BEGIN -->
Run reserving-class file mutations that a Client PC would otherwise perform one SMB round trip at a time on the ArcRho Server host over HTTP. Like [`workspace_reads`](workspace_reads.md) this is a transport, not a domain of its own: the canonical `app_server` service function still owns the operation and runs unchanged either locally or inside the machine-wide ArcRho Gateway.

It is a separate registry and a separate route from workspace reads because the reasoning that makes a read safe does not carry over. A read is a pure function of the workspace, so an uncertain answer may simply be asked again or answered locally instead; a mutation may not. Every registered mutation kind must be **idempotent** — running it twice must leave the same end state as running it once — which is what lets this transport, like the read one, keep no durable receipt. A mutation that cannot make that promise belongs on the hosted-save path (`arcrho_hosted_save_http_contract`), which buys durability with a request file, an Engine claim, and a receipt.
<!-- MANUAL:END -->

## Entry Points
<!-- MANUAL:BEGIN -->
No new browser-facing route. This existing route selects the transport per request through `workspace_mutation_client.run_workspace_mutation`:

| Route | Mutation kind | Service |
| --- | --- | --- |
| `POST /datasets/cached/delete` | `cached_dataset_delete` | `dataset_service.delete_cached_datasets` |
| `POST /dfm/rpc-bridge/sync` | `dfm_rpc_bridge_sync` | `dfm_rpc_bridge_service.hosted_send_sync_request` |
| `POST /dfm/rpc-bridge/keep-local` | `dfm_rpc_bridge_keep_local` | `dfm_rpc_bridge_service.hosted_keep_local` |
| `POST /dfm/rpc-bridge/cleanup` | `dfm_rpc_bridge_cleanup` | `dfm_rpc_bridge_service.hosted_cleanup_tmp` |
| `POST /dfm/rpc-bridge/update-remote` | `dfm_rpc_bridge_update_remote` | `dfm_rpc_bridge_service.hosted_update_remote` |
| (no route; the ResQ sync and export macros call `run_workspace_mutation` directly through `arcrho_api.resq_sync_queue.submit_sync_request`) | `resq_sync_request_publish` | `resq_sync_queue_service.publish_resq_sync_request` |
| (no route; the two ResQ import macros call `run_workspace_mutation` directly through `arcrho_api.resq_import_backup.back_up_reserving_class`) | `resq_import_backup` | `resq_import_backup_service.back_up_reserving_class_for_import` |

The ResQ sync-queue publish is the request file the Sync and Export Reserving
Class with ResQ macros hand to a ResQ-connected Bridge worker. The payload and
the on-disk write are `arcrho_api.resq_sync_queue`'s own; the service adds only
the place it is written from and stamps the acting user's login as the
request's `UserName`. It is idempotent by request id — an id that already has
a request or a status file is returned as `resumed` — and the macro's polling
of that request's status is the hosted `bridge_worker_liveness` read, so
inside the app neither half of the exchange crosses the share.

The ResQ import backup is the copy both import macros take of the reserving
class they are about to rewrite. It is the largest of these kinds by file
count -- every method, every sidecar a person could have edited, every data
file those sidecars name, and the class index -- so from a Client PC it was one
SMB round trip per file, plus a class-folder lookup that can read one index per
reserving class of the project. Hosted, the whole copy is local disk and the
macro pays one request. The copy itself, the folder layout, the retention rule
and what it leaves out are all `arcrho_api.resq_import_backup`'s; the service
adds only the place it runs and stamps the acting user's login into the
manifest. It is idempotent by backup id: the macro owns the id, and an id whose
copy this host already finished -- which its `backup.json` records -- is
reported as it stands rather than copied again under a second folder. A copy
that died part way leaves no manifest and is never presented as a restore
point. Because a backup that cannot be taken never stops an import, the macro
turns a transport failure into a warning rather than an error, and words an
unconfirmed outcome as unknown rather than as "no restore point".

The RPC-bridge kinds are the one family whose work is not finished when the
service returns from local disk: `sync` and `update-remote` publish a request
file and then wait for the ArcRho Bridge to answer. The Bridge runs on the
server host, so hosting these puts both halves of that exchange on local disk,
where the wait is a file-system event rather than a poll that must write and
delete a probe file to defeat the SMB redirector's cached "not found". Their
keyword arguments are the route schema's own fields and the service rebuilds
its request model from them, so pydantic stays the only validator; the wait a
caller may ask the Gateway to hold a thread for is clamped by
`clamp_rpc_bridge_wait`.

Gateway side: `POST /api/workspace-mutations` on the Gateway (`arcrho_workspace_mutation_contract.WORKSPACE_MUTATION_PATH`), authenticated with the same per-user HMAC headers as hosted saves and workspace reads. `GET /api/capabilities` advertises `workspace_mutation_kinds`.
<!-- MANUAL:END -->

## Key Files
<!-- MANUAL:BEGIN -->
- `python-api/src/arcrho_workspace_mutation_contract.py` - The canonical `WORKSPACE_MUTATION_KINDS` registry (kind → service module, function, required/optional keyword arguments, and which of them are name lists), request validation, route path, and timeout.
- `app_server/services/workspace_mutation_client.py` - Client transport selection and the no-fallback-after-acceptance rule. It reuses the read transport's signing, capability probe cache, `post_signed_json`, and path rebasing rather than repeating them.
- `server-components/src/arcrho_gateway/workspace_mutations.py` - Server-side executor: authenticates, validates against the registry, imports the bundled service, runs it under `acting_identity`, and maps a service `HTTPException` to the same status while preserving a structured refusal detail.
- `server-components/src/arcrho_gateway/main.py` - Route dispatch and the capability field; the handler is the one `_handle_hosted_execution` shared with reads and calculations.
- `server-components/src/arcrho_gateway/build_exe.py` - Registered mutation service modules join the hidden-import list and the pre-build import probe automatically.
<!-- MANUAL:END -->

## Data/State/Caches
<!-- MANUAL:BEGIN -->
- Transport selection matches workspace reads exactly: a process running with `ARCRHO_RUNTIME_SERVER_ROOT` set (Engine, Bridge, Gateway) always runs locally so the gateway can never route back to itself; otherwise the local `%APPDATA%\ArcRho\arcrho_gateway.json` credential must be enabled and the gateway's `/api/capabilities` must list the mutation kind.
- Fallback rule, and the one real difference from reads: the client falls back to the mapped drive **only** when the failure proves the server never acted (`GatewayTransportFailure.accepted` is false — unreachable, authentication refused, an older gateway without the route). A timeout or a connection lost after the request was sent surfaces as `504` telling the user to refresh and look, because a local re-run would be reasoning about a workspace the server has already changed.
- A refusal raised by the hosted service itself is recognized by the `X-ArcRho-Workspace-Root` header the gateway sets whenever the operation ran, and passes through with the status the local path would have raised.
- Structured refusals survive the wire. `POST /datasets/cached/delete` answers `409` with an object (`error`, `message`, `blocked_datasets`) that Project Instance renders as the dependents window, so the transport preserves a mapping `detail` instead of flattening it to text; only its free text is redacted for server paths.
- Path rebasing is the read transport's: every string in the response that starts with the server's workspace root is rewritten onto this PC's own root, so `deleted_files[].path` and the returned `index.folder_paths` look exactly as a local delete would have produced them.
- Identity: the request carries the enrolled `UserName` and the client's resolved display name, and the gateway binds `user_identity_service.acting_identity` around the mutation so anything it stamps on disk names the person who asked.
- No idempotency receipt: registered kinds are idempotent by contract, so a lost response is safe to ask about again rather than needing a replay record.
- Diagnostics: one record per mutation in `client_read_latency.jsonl`, keyed `read_kind: "mutation:<kind>"`, with the same `transport` / `reason` / timing fields the reads use.
<!-- MANUAL:END -->

## Common Change Tasks
<!-- MANUAL:BEGIN -->
1. Move another mutation to the server: confirm it is idempotent, add one `WorkspaceMutationKind` entry naming the service function and its keyword arguments (listing any list-valued argument in `list_args`), then wrap the route's service call in `workspace_mutation_client.run_workspace_mutation(...)`. `test_workspace_mutations.py` fails if the registry names an argument the function lacks or omits one it requires; the gateway build validates the import graph and bundles the module automatically. Rebuild and redeploy the gateway.
2. If the operation is not idempotent, do not add it here — put it on the hosted-save path instead, which owns receipts, Engine claims, and the reserving-class lease.
3. A refusal the page acts on rather than merely displays should raise a mapping `detail` from the service. Both transports deliver the same object, so the page needs one shape; keep any server path out of the structured fields, because only free text is redacted.
<!-- MANUAL:END -->

## Known Risks
<!-- MANUAL:BEGIN -->
- The idempotence requirement is a contract promise, not something the transport can verify. A kind registered here that quietly stops being idempotent would turn a lost response into a silently wrong end state.
- The gateway is one process for the whole fleet, and a mutation holds a handler thread for its whole run; a large delete competes with hosted saves and reads in the same process.
- The pilot transport is plain HTTP with HMAC, so dataset names travel unencrypted on the internal network. TLS is the first gate before broader rollout.
<!-- MANUAL:END -->
