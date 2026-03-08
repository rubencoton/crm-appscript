# AutoSync GitHub

## What this does
- Scans all Git repos inside `C:\Users\elrub\Desktop\CARPETA CODEX`.
- In `backup` mode, skips `clasp pull` to avoid churn and only backs up local development changes.
- In `full` mode, runs `clasp pull` for projects that contain `.clasp.json`.
- Auto-fixes `safe.directory` if Git blocks a repo by ownership checks.
- Runs `git pull --rebase --autostash` (if upstream exists).
- If there are changes, auto-commits and pushes to GitHub.
- Adds local exclude rules in `.git\info\exclude` for `node_modules/`, logs, and lock files.
- Writes logs to `06_AUTOMATIZACION\logs`.

## Recommended interval
- `10 minutes`.

## Run manually (safe backup mode)
```powershell
cd "C:\Users\elrub\Desktop\CARPETA CODEX\06_AUTOMATIZACION"
powershell -ExecutionPolicy Bypass -File .\sync-all.ps1 -Mode backup -WorkspaceRoot "C:\Users\elrub\Desktop\CARPETA CODEX"
```

## Run manually (full mode with clasp)
```powershell
cd "C:\Users\elrub\Desktop\CARPETA CODEX\06_AUTOMATIZACION"
powershell -ExecutionPolicy Bypass -File .\sync-all.ps1 -Mode full -WorkspaceRoot "C:\Users\elrub\Desktop\CARPETA CODEX"
```

## Active scheduled task
- `\Codex-AutoSync-GitHub`
- Triggers: every 10 minutes + at logon

## Remove scheduled task
```powershell
schtasks /Delete /TN "Codex-AutoSync-GitHub" /F
```
