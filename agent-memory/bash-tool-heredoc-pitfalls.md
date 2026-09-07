---
name: bash-tool-heredoc-pitfalls
description: "In this Windows Git-Bash tool, heredoc bodies are not literal — backticks abort the command and `\\\\` collapses to `\\`; write scripts with the Write tool and run them by path"
metadata: 
  node_type: memory
  type: project
  originSessionId: be39cc59-f1e3-4e66-8f1b-060ea472932d
  modified: 2026-09-07T16:06:56.601Z
---

On this Client PC the Bash tool's `<<'EOF'` heredocs are **not** treated as literal text even with a quoted delimiter (observed 2026-08-17):

- Any backtick in the body (JS template literals, Markdown code spans in a doc-editing script) makes bash fail with ``unexpected EOF while looking for matching `''`` — the whole command runs nothing.
- Doubled backslashes are collapsed: a Python line `"Z:\\A\\New.xlsx"` inside the heredoc reaches Python as `"Z:\A\New.xlsx"` and raises `unicodeescape` errors, or silently writes single-backslash text into files.

**Why:** the command string is pre-processed before Git Bash sees it, so heredocs only work for bodies free of backticks and escaped backslashes.

The backslash collapse also silently produces *parseable but broken* JS, so a syntax check can pass while the browser fails. On 2026-09-07 a `cat >> open_path.js <<'EOF'` block landed `` `${folder}\` `` instead of `` `${folder}\\` ``; `node --check` still passed (the stray backtick opened a template that closed further down), and the only symptom was every Dataset Viewer page loading blank with `Uncaught SyntaxError` in the renderer. Catching it needed the [[electron-ui-screenshot-check]] probe with a `console-message` listener pointed at the running app's page URL. This applies even while auto mode asks for Bash-first edits: code with backticks or backslashes still goes through Write/Edit.

**How to apply:** for anything longer than a few plain lines — especially JS/CSS/Markdown edits or scripts with Windows paths — write the script or content with the Write tool into the scratchpad and run it by path (`py -3.10 <scratchpad>/script.py`), or use the Edit tool directly. Related: [[python-test-runner]], [[frontend-node-test-suite]].

A rewrite script must also keep line endings: the working copies here are CRLF (`git ls-files --eol` shows `i/lf w/crlf`), so `Path.read_text()`/`write_text()` silently turns a whole file to LF (observed 2026-08-27). Open with `newline=""`, convert the search/replace strings to the file's own EOL, and write with `newline=""` — or restore CRLF afterwards with a bytes replace.

`sed -i` in this Bash tool does the same (observed 2026-09-07 on the Server PC while bumping `?v=` stamps): every file it rewrites comes out LF-only, and `git stash` then warns "LF will be replaced by CRLF" for each one. A `git stash` + `git stash pop` cycle re-checks the files out and restores CRLF, and `git diff` never shows the ending change because the index is LF anyway, but a deploy patch or a test that reads bytes sees the LF file in between. Prefer the Edit tool for one-line pin bumps, or run `git ls-files --eol` on the touched files afterwards and fix any `w/lf`.

A second trap in the same rewrite scripts (2026-09-07): `text = path.read_text()` has already translated every CRLF to LF, so a probe such as `nl = CRLF if CRLF in text else LF` is always false and the file is written back LF even though the script "checked". Detect the ending on the bytes (`path.read_bytes()` containing CR+LF) or open with `newline=""`. Do not trust a `grep -c` for a trailing CR in this Bash tool as a CRLF check either: it reported every line as CRLF on files that were pure LF. `git ls-files --eol` and a bytes count are the reliable checks.
