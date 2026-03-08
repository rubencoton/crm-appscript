# AutoSync GitHub

## What this does
- Scans all Git repos inside `C:\Users\elrub\Desktop\CARPETA CODEX`.
- Runs `clasp pull` in each folder that contains `.clasp.json`.
- Auto-fixes `safe.directory` if Git blocks a repo by ownership safety checks.
- Runs `git pull --rebase --autostash` (if upstream exists).
- If there are changes, auto-commits and pushes to GitHub.
- Adds local exclude rules for `node_modules/` to avoid giant accidental commits.
- Writes logs to `06_AUTOMATIZACION\logs`.

## Recommended interval
- `10 minutes`.

## One-time setup
```powershell
cd "C:\Users\elrub\Desktop\CARPETA CODEX\06_AUTOMATIZACION"
powershell -ExecutionPolicy Bypass -File .\install-autosync-task.ps1 -IntervalMinutes 10
```

## Run manually now
```powershell
cd "C:\Users\elrub\Desktop\CARPETA CODEX\06_AUTOMATIZACION"
powershell -ExecutionPolicy Bypass -File .\sync-all.ps1 -Mode backup
```

## Remove scheduled task
```powershell
cd "C:\Users\elrub\Desktop\CARPETA CODEX\06_AUTOMATIZACION"
powershell -ExecutionPolicy Bypass -File .\remove-autosync-task.ps1
```
