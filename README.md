# Propuestas — Pliegos tipo vías Colombia
## Resolución 465/2024 · Colombia Compra Eficiente · Versión 4

Sistema para generar propuestas en licitaciones públicas de infraestructura de transporte.

## Cómo subir a Railway

1. Sube esta carpeta a GitHub (ver instrucciones abajo)
2. En Railway: New Project → Deploy from GitHub repo
3. Selecciona el repositorio
4. En Variables de entorno agrega: ANTHROPIC_API_KEY = tu_clave_api
5. Railway despliega automáticamente

## Subir a GitHub (primera vez)

1. Ve a github.com y crea un repositorio nuevo llamado "propuestas-vias"
2. En tu computador instala Git: git-scm.com
3. Abre la terminal en la carpeta del proyecto y ejecuta:
   git init
   git add .
   git commit -m "primera version"
   git remote add origin https://github.com/TU_USUARIO/propuestas-vias.git
   git push -u origin main

## Módulos

- Base de datos de empresas: registra empresas, sube PDFs, la IA extrae datos
- Nuevo proceso: sube el pliego, la IA detecta qué formatos aplican
- Verificar y generar: cruza indicadores financieros y descarga todos los Word
