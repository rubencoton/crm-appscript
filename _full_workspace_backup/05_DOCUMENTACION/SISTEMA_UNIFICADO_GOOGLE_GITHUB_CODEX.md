# SISTEMA UNIFICADO GOOGLE + GITHUB + CODEX (2 EQUIPOS)

Fecha de verificacion: 2026-03-08

## Objetivo
Tener sobremesa y portatil sincronizados sin conflictos entre:
- Google Apps Script
- GitHub
- Repositorio local en Codex
- Copia de seguridad en Google Drive (solo backup)

## Regla clave
No usar Google Drive como carpeta de trabajo viva del codigo.
Google Drive solo para backup one-way.

## Fuente de verdad
1. GitHub (repo privado) es la fuente de verdad del codigo.
2. Cada equipo trabaja en su clon local.
3. Apps Script se sincroniza con clasp desde el repo local.
4. Drive guarda snapshots de respaldo, nunca sync bidireccional del repo.

## Estado actual detectado
- Repo local principal:
  - C:\Users\elrub\Desktop\CARPETA CODEX\01_PROYECTOS\festivales-github
- Remote GitHub:
  - https://github.com/rubencoton/crm-appscript.git
- Cuenta clasp autenticada en este equipo:
  - manager@rubencoton.com
- ScriptId principal (repo raiz):
  - 1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea

## Scripts creados
- Setup de equipo:
  - C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\setup_equipo_sync.ps1
- Sync diario:
  - C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\sync_git_gas.ps1

## Flujo diario recomendado (en cada equipo)
1. Bajar todo antes de trabajar:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\sync_git_gas.ps1" -Action down
```

2. Ver estado:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\sync_git_gas.ps1" -Action status
```

3. Subir todo al terminar:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\sync_git_gas.ps1" -Action up -CommitMessage "tu mensaje"
```

## Backup a Google Drive (opcional)
Ejemplo:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\elrub\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\sync_git_gas.ps1" -Action backup -DriveBackupPath "D:\Google Drive\BACKUPS_CODIGO"
```

## Alta del portatil (solo 1 vez)
Ejecutar en el portatil:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\Users\TU_USUARIO\Desktop\CARPETA CODEX\03_SCRIPTS_UTILIDAD\setup_equipo_sync.ps1"
```

Ese script:
- Clona el repo
- Ejecuta npm install
- Verifica login clasp
- Verifica estado del proyecto Apps Script

## Conversaciones de Codex en ambos equipos
Las conversaciones no se sincronizan por carpeta local.
Se sincronizan por cuenta de la app/plataforma.

Para ver hilos en ambos dispositivos:
1. Inicia sesion con la misma cuenta en los dos equipos.
2. Usa el mismo workspace cuando quieras continuidad de contexto.
3. Si un hilo es critico, deja resumen tecnico en 05_DOCUMENTACION para continuidad operativa.

## Riesgos evitados con esta arquitectura
- Conflictos por doble escritura (Drive + Git)
- Perdida silenciosa de archivos por sync bidireccional
- Divergencia entre codigo en Apps Script y codigo local

## Checklist rapido
- [ ] Repo local fuera de carpeta sincronizada de Drive
- [ ] git pull + clasp pull antes de editar
- [ ] clasp push + git commit/push al terminar
- [ ] Backup puntual a Drive (no trabajo en vivo)
- [ ] Misma cuenta Codex en ambos equipos
