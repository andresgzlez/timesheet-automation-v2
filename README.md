# Automatizacion de Timesheets

App web para procesar timesheets automaticamente (ver detalle del proyecto
en la conversacion con Claude / el documento de plan).

## Como desplegarlo (GitHub + Railway)

### 1. Crear cuenta de GitHub (2 min, gratis)
1. Entra a https://github.com/signup
2. Pon tu correo, una contrasena, y un nombre de usuario.
3. Confirma el correo que te llega.

### 2. Subir este codigo a un repositorio
1. En GitHub, boton verde **"New"** (o el `+` arriba a la derecha -> "New repository").
2. Nombre: `timesheet-automation` (o el que quieras). Puede quedar **privado**.
3. NO marques "Add a README" (ya trae uno).
4. Click **Create repository**.
5. GitHub te muestra unos comandos bajo "...or push an existing repository
   from the command line" -- son los que se usan para subir esta carpeta.
   (Si no sabes correrlos, Claude puede hacerlo por ti si le compartes el
   link del repositorio vacio que acabas de crear.)

### 3. Conectar Railway
1. Entra a https://railway.com y crea cuenta (puedes entrar directo con tu
   cuenta de GitHub, mas rapido).
2. **New Project -> Deploy from GitHub repo** -> elige el repositorio que
   subiste.
3. Railway detecta que es Python solo, y usa el `Procfile` para saber como
   arrancarlo. No hay que tocar nada mas.
4. Cuando termine de construirse (1-2 min), Railway te da un link publico
   (Settings -> Networking -> Generate Domain).

Listo -- ese link es la app, funcionando para cualquiera que lo tenga,
desde cualquier dispositivo.

## Limitacion conocida (v1)

Project Daisy llega como archivo `.xls` (formato viejo de Excel) y este
programa lo convierte con LibreOffice antes de leerlo. El servidor de
Railway por defecto **no trae LibreOffice instalado**, asi que por ahora,
para ese proyecto especifico, hay que abrir el archivo `.xls` en Excel y
guardarlo como `.xlsx` antes de subirlo a la web app. Se puede arreglar
agregando LibreOffice a la configuracion de Railway mas adelante si hace
falta.

## Estructura del proyecto

- `engine/` -- logica central: matching de nombres, calculo de REG/OT,
  escritura de Excel preservando formulas, deteccion automatica de layout.
- `profiles/` -- configuracion por proyecto (solo los que no se detectan
  solos, como HCA Palms West).
- `app/` -- la app web (FastAPI) y sus plantillas HTML.
