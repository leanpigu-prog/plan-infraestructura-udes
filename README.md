# Plan de Infraestructura Física UDES

Herramienta interna para la planeación y el seguimiento de los 19 proyectos transversales de infraestructura física identificados en el `Plan Maestro de Desarrollo Físico y Tecnológico (PDIFyT) UDES 2026` (Tabla 18, pág. 47-48), organizados en 7 programas (categorías funcionales) y 3 campus (Bucaramanga, Cúcuta, Valledupar).

- **Frontend:** `index.html` (estático, publicado con GitHub Pages).
- **Backend:** Google Apps Script + Google Sheets (`apps-script/Code.gs`).

## Puesta en marcha (una sola vez)

### 1. Crear el Google Sheet y el backend

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja de cálculo nueva llamada, por ejemplo, **"Plan Infraestructura UDES - Base de datos"**.
2. Abre **Extensiones → Apps Script**.
3. Borra el contenido de `Code.gs` que aparece por defecto y pega ahí el contenido de `apps-script/Code.gs` de este repo.
4. Guarda el proyecto de Apps Script (ícono de disquete).
5. En el selector de funciones (arriba, junto a "Depurar"), elige `setup` y presiona **Ejecutar**.
   - La primera vez te pedirá autorizar permisos: acepta con tu cuenta institucional.
   - Esto crea 3 pestañas en tu Sheet: `Proyectos` (con los 19 proyectos precargados), `Usuarios` y `Historial_Seguimiento`.
6. Abre el Sheet y ve a la pestaña **Usuarios**: cambia las claves de ejemplo (`CAMBIAR-CLAVE-1`, etc.) por claves reales para cada campus y para el usuario `admin`. **No dejes las claves de ejemplo.**

### 2. Desplegar el backend como Web App

1. En el editor de Apps Script, botón **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. "Ejecutar como": **Yo (tu cuenta)**.
4. "Quién tiene acceso": **Cualquier usuario** (necesario para que el sitio público pueda llamarlo; el acceso real a los datos lo controla el login usuario/clave dentro de la app, no este permiso).
5. Implementar → copia la **URL de la aplicación web** que te entrega (termina en `/exec`).

### 3. Conectar el frontend con el backend

1. Abre `index.html` en este repo.
2. Busca la línea:
   ```js
   const API_URL = 'PON_AQUI_LA_URL_DEL_WEB_APP';
   ```
3. Reemplaza `'PON_AQUI_LA_URL_DEL_WEB_APP'` por la URL que copiaste en el paso anterior.
4. Guarda, haz commit y push. GitHub Pages se actualiza solo en 1-2 minutos.

## Usuarios por defecto (cambiar claves antes de usar)

| Usuario | Campus | Rol |
|---|---|---|
| bucaramanga | Bucaramanga | responsable |
| cucuta | Cúcuta | responsable |
| valledupar | Valledupar | responsable |
| admin | Todos | admin (vista consolidada) |

## Notas de seguridad

Esta herramienta usa un login simple (usuario/clave en texto plano en el Sheet) pensado para un grupo pequeño y conocido de responsables institucionales — **no** es apta para exponer datos sensibles a público general ni resiste intentos de fuerza bruta. Si el costo estimado o los datos de los proyectos se consideran confidenciales, evalúa restringir el acceso al repositorio/Sheet o añadir un mecanismo de autenticación más robusto antes de compartir la URL ampliamente.

## Cada responsable de campus puede

- Ver únicamente los proyectos de su campus.
- Editar: costo estimado, fuente de financiación, fechas de inicio/fin, fase actual, responsable, estado, prioridad y % de avance.
- Cada edición queda registrada en la pestaña `Historial_Seguimiento` del Sheet (trazabilidad en el tiempo).

## El usuario `admin` puede

- Ver los 19 proyectos de los 3 campus.
- Ver el consolidado por programa × campus, alertas de proyectos atrasados y el costo total estimado.
- Editar cualquier proyecto.
