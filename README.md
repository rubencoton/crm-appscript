# CRM Festivales - Apps Script + GitHub

Este repo esta conectado al proyecto de Google Apps Script:
- scriptId: `1OGuPezQ26BFvaLRiy-IYIotGpmVu_Z_b9Mi8tCiprIz8zB4DgqmMc5Ea`
- archivo de enlace: `.clasp.json`

## Reglas de sincronizacion
- Fuente de verdad: GitHub
- Carpeta local de trabajo: este repo
- Apps Script: se sincroniza con `clasp`
- Google Drive: solo backup (no usar como carpeta activa del repo)

## Comandos base
- `npm run gas:login` -> autenticar clasp en el equipo
- `npm run gas:status` -> ver estado Apps Script
- `npm run gas:pull` -> traer cambios desde Google
- `npm run gas:push` -> subir cambios a Google
- `npm run gas:open` -> abrir editor Apps Script

## Comandos de sincronizacion
- `npm run sync:down` -> `git pull` + `clasp pull`
- `npm run sync:status` -> estado de clasp + estado git
- `npm run sync:up` -> `clasp push` + stage de git

## Flujo diario recomendado
1. `npm run sync:down`
2. Editar codigo
3. Probar
4. `npm run sync:up`
5. `git commit -m "mensaje"`
6. `git push`

## Nota multi equipo
Si usas sobremesa y portatil:
- ambos con la misma URL de repo
- ambos con `npm install`
- ambos con `npm run gas:login`
