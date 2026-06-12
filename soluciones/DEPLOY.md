# 🚀 Guía de Despliegue en Railway

## Requisitos previos
- Cuenta gratuita en [github.com](https://github.com) (si no tienes una)
- Cuenta gratuita en [railway.app](https://railway.app)

---

## PASO 1 — Subir el proyecto a GitHub

1. Ve a **github.com** → botón verde **"New"** → crea un repositorio
   - Nombre: `soluciones-logisticas`
   - Visibilidad: **Private** (recomendado)
   - Sin README ni .gitignore (ya incluidos)

2. En tu computadora, abre la terminal dentro de la carpeta `soluciones/`:

```bash
git init
git add .
git commit -m "primer deploy"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/soluciones-logisticas.git
git push -u origin main
```

---

## PASO 2 — Desplegar en Railway

1. Ve a [railway.app](https://railway.app) → **"Start a New Project"**
2. Elige **"Deploy from GitHub repo"**
3. Conecta tu cuenta de GitHub y selecciona `soluciones-logisticas`
4. Railway detecta automáticamente que es Flask y despliega

---

## PASO 3 — Configurar variables de entorno en Railway

En el panel de Railway → tu proyecto → pestaña **"Variables"**, agrega:

| Variable       | Valor                              |
|----------------|------------------------------------|
| `SECRET_KEY`   | `una-clave-segura-larga-aleatoria` |
| `FLASK_ENV`    | `production`                       |

> Para generar una clave segura puedes usar:
> ```python
> python -c "import secrets; print(secrets.token_hex(32))"
> ```

---

## PASO 4 — Obtener tu URL pública

En Railway → tu proyecto → pestaña **"Settings"** → sección **"Domains"**:

- Haz clic en **"Generate Domain"**
- Obtendrás algo como: `https://soluciones-logisticas-production.up.railway.app`

¡Esa es tu liga pública! Compártela con tu equipo.

---

## ⚠️ Nota importante sobre la base de datos

Railway usa un **sistema de archivos efímero**: la base de datos SQLite
se resetea cada vez que se redespliega la app.

**Para producción real se recomienda agregar PostgreSQL:**

1. En Railway → **"New Service"** → **"Database"** → **PostgreSQL**
2. Railway agrega automáticamente la variable `DATABASE_URL`
3. Instala el driver: agrega `psycopg2-binary==2.9.9` a `requirements.txt`

Si el volumen de usuarios es bajo y no te preocupa perder datos al redesplegar,
SQLite funciona bien para empezar.

---

## 📋 Estructura final del proyecto

```
soluciones/
├── app.py
├── models.py
├── requirements.txt     ← incluye gunicorn
├── Procfile             ← le dice a Railway cómo arrancar
├── railway.json         ← configuración de Railway
├── .gitignore
├── templates/
│   ├── base.html
│   ├── login.html
│   ├── dashboard.html
│   ├── nueva_solicitud.html
│   ├── solicitudes.html
│   └── detalle_solicitud.html
└── static/
    ├── style.css
    └── app.js
```

---

## 👤 Usuarios demo (primer arranque)

| Usuario  | Contraseña | Rol              |
|----------|------------|------------------|
| lider    | lider123   | Líder Soluciones |
| gerardo  | hunter123  | Hunter           |
| laura    | hunter123  | Hunter           |
| juan     | sol123     | Soluciones       |
| maria    | sol123     | Soluciones       |

> **Cambia las contraseñas** desde el código antes de dar acceso real al equipo.
