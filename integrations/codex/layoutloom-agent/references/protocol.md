# LayoutLoom Agent protocol

## Request schema

Save requests as UTF-8 JSON objects no larger than 1 MiB:

```json
{
  "schema_version": "1.0",
  "request_id": "optional-stable-id",
  "operation": "pdf.to_word",
  "inputs": ["D:\\documents\\paper.pdf"],
  "output_dir": "D:\\documents\\output",
  "parameters": {
    "mode": "hybrid",
    "column_layout": "auto"
  },
  "options": {
    "expand_globs": false
  }
}
```

Relative paths are resolved against the request file's directory. Prefer absolute paths for auditability. Unknown fields, duplicate JSON keys, unknown parameters, invalid choices, missing files, and unsupported extensions fail validation.

Passwords may appear inside `parameters`, but the CLI never returns their values. Do not pass secrets with `-p` or place them in logs.

## Commands

- `protocol`: return protocol and application versions.
- `catalog [--query TEXT] [--group NAME] [--full] [--probe]`: discover operations. Capability probing is opt-in because external engine checks may be slow.
- `describe OPERATION`: return the complete parameter and input schema and probe that operation's current capability.
- `validate --request FILE`: check request structure, paths, parameters, operation capability, and optional path roots without creating the output directory.
- `run --request FILE --format jsonl`: execute through LayoutLoom's supervised `TaskRunner`.
- `quick-run OPERATION INPUT... --output-dir DIR --format jsonl`: build the request in memory and execute it in one process. Use this by default for known, non-destructive tasks; it performs the same runtime validation as `run`.
- `install-skill [--force]`: install this Skill into the current user's Codex skills directory.

Use `--request -` to read the request from UTF-8 stdin. Use repeated `--allow-root PATH` host arguments to restrict all input, output, and path-type parameters to authorized locations.

`quick-run` accepts repeated non-sensitive `--param key=value` options. Password parameters are rejected on the command line and must be supplied through a UTF-8 JSON object with `--params-file`.

## Fast Office conversion policy

For the WPS-optimized LayoutLoom installation, invoke `word.to_pdf`, `excel.to_pdf`, and `ppt.to_pdf` with `--param engine=wps` unless the user selected another engine or WPS has already reported that it is unavailable. Do not probe a same-named output before the call; LayoutLoom creates a unique path instead of overwriting an existing file.

Run one Office `quick-run` at a time and treat its JSONL stream as the source of truth. Wait at most 90 seconds for a final event. If the limit is reached, send one graceful interrupt and wait up to 15 more seconds for cancellation cleanup. A retry may start only after the previous wrapper process has exited, and it must use a new independent process/tool call rather than the active terminal session. Do not kill generic WPS or Microsoft Office processes.

## JSONL events

Every stdout line is one JSON object with a monotonically increasing `seq`:

- `accepted`: request accepted; no result yet.
- `progress`: `fraction`, `percent`, and a user-readable stage message.
- `cancel_requested`: graceful cancellation has started.
- `result`: final outcome and all output, warning, completed, failed, and cancelled input lists.
- `error`: final stable error code and message. Cancellation may include `partial_result`.

Do not treat `accepted` or `progress=1.0` as completion. Wait for exactly one final `result` or `error` event.

## Exit codes

- `0`: success.
- `2`: invalid JSON, operation, parameter, path, input, or CLI usage.
- `3`: required local engine unavailable.
- `4`: processing failed with a handled LayoutLoom error.
- `5`: partial batch success; inspect both `outputs` and `failed_inputs`.
- `70`: unexpected internal error.
- `130`: cancelled; inspect `partial_result` for preserved outputs.

## Safety behavior

LayoutLoom validates input count and extensions, rejects undeclared operations and parameters, locks the output directory, requires declared outputs to exist, blocks outputs outside the selected directory unless an operation explicitly permits them, and cleans temporary or empty artifacts after cancellation. Independent batch operations continue after one input fails. Ordered combined operations preserve their input order.

Agent mode blocks `image.rename` with `move=true` unless the host supplies `--allow-source-mutation`. The request JSON cannot grant that permission to itself.
