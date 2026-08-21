---
name: layoutloom-agent
description: Use LayoutLoom's local JSON CLI to run PDF, Word, Excel, PowerPoint, image, and video processing operations. Use when an agent must quickly process local files with LayoutLoom, discover operation IDs or parameters, validate complex or batch work, monitor progress, handle partial results, or install/check the LayoutLoom bridge.
---

# LayoutLoom Agent

Use the bundled wrapper so source checkouts and portable Windows builds share one command surface.

## Default to the one-call fast path

For a clear, authorized, non-destructive task with a known operation ID and output directory, call `quick-run` once. Do not first call `protocol`, `catalog`, `describe`, or `validate`. Do not ask the user to reconfirm an ordinary “另存为” task when the input, operation, and output location are already clear.

```powershell
python scripts/layoutloom_agent.py quick-run word.to_pdf `
  "C:\资料\合同.docx" `
  --output-dir "C:\资料\输出" `
  --param engine=wps `
  --allow-root "C:\资料" `
  --format jsonl
```

`quick-run` builds an in-memory request and performs the same schema, input, parameter, path, engine, cancellation, output-lock, and final-file checks as `run`. It skips only redundant discovery and a separate dry-run process.

Common stable operation IDs:

- `word.to_pdf`, `excel.to_pdf`, `ppt.to_pdf`
- `pdf.to_word`, `pdf.merge`, `pdf.split`, `pdf.compress`, `pdf.compress_lossy`
- `image.convert`, `image.resize`, `image.crop`, `image.rotate`, `image.compress`
- `video.transcode`, `video.compress`, `video.trim`, `video.extract_audio`
- `video.extract_slides_ppt`

Use defaults when they match the request. Pass known non-sensitive values with repeated `--param key=value`.

For `word.to_pdf`, `excel.to_pdf`, and `ppt.to_pdf` on this WPS-optimized installation, explicitly pass `--param engine=wps` unless the user requested another engine or WPS has already returned an unavailable-engine error. Do not leave these fast conversions on automatic engine selection: fallback probing can start Microsoft Office and turn a seconds-long conversion into a long wait. Never silently replace an engine explicitly selected by the user.

Do not pre-check whether a same-named output already exists. LayoutLoom generates a unique output path and does not overwrite an existing result. After the final event, verify only the output paths reported by LayoutLoom.

Read every JSONL line until one final `result` or `error` event appears. Verify each reported output exists and is non-empty before reporting completion.

### Office fast-path lifetime

- Start exactly one `quick-run` process for an Office-to-PDF request. Its JSONL stream is the authoritative status; do not add file-timestamp polling, generic Office-process inspection, or a second conversion while it is active.
- Allow up to 90 seconds for a final event. If none arrives, send one graceful interrupt, then allow up to 15 seconds for LayoutLoom to emit its cancellation result and clean temporary files.
- Never type or launch a retry inside the still-running terminal session. Retry only after the first wrapper process has exited and released its output lock, using a new independent tool/process call.
- Retry at most once for a transient failure. Keep `engine=wps` unless the final error explicitly says WPS is unavailable; report that dependency error instead of silently beginning a long fallback chain.
- If graceful cleanup does not finish, stop only the exact LayoutLoom wrapper process owned by this call. Do not terminate arbitrary `wps.exe`, `WINWORD.EXE`, or user Office processes.

## Use discovery only when needed

Use the longer workflow when the operation ID or parameter schema is genuinely unknown:

1. Call `protocol` only for first-time bridge diagnosis, a suspected version change, or an incompatible-command error.
2. Call `catalog --query ...` only to find an unknown operation ID.
3. Call `describe OPERATION` only before choosing unfamiliar or non-default parameters.
4. Create a UTF-8 JSON request.
5. Call `validate` separately for complex batches, source-moving work, sensitive parameters, multiple auxiliary path assets, or uncertain external-engine selection.
6. Call `run --request ... --format jsonl`.

Do not repeat discovery during the same task after the application and protocol versions are known.

## Sensitive parameters

Do not put passwords in command-line arguments. `quick-run --param` rejects password fields. Put sensitive values in a temporary UTF-8 JSON parameter object and use `--params-file`; remove the temporary file after the task when practical.

For a full request, put secrets only in the request JSON, never echo its contents, and remove it after use when practical.

## Guardrails

- Use only stable IDs listed above or returned by `catalog`/`describe`.
- Prefer one or more `--allow-root` arguments covering every authorized input, output, and auxiliary asset root.
- Do not use `--allow-source-mutation` unless the user explicitly authorizes moving originals. `image.rename move=true` is blocked otherwise.
- Do not secretly change a requested engine when it is unavailable. Report the missing WPS, Microsoft Office, LibreOffice, FFmpeg, Poppler, or GPU dependency.
- Treat `video.repair_slides_ppt` as GUI-assisted and run it only with a repair plan created by LayoutLoom's repair workspace.
- On cancellation, send one graceful interrupt and allow cleanup to finish before considering a forced stop of the exact task-owned wrapper process.
- For exit code `5`, preserve successful outputs and report failed inputs. For exit code `130`, report cancellation and any preserved partial outputs.

Read [references/protocol.md](references/protocol.md) when constructing full requests or interpreting batch, cancellation, and exit-code behavior.
