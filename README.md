# Travesías — convocatoria

Página de la convocatoria de fotografía y video de **Travesías**
(Universidad de los Andes · Antropología × Estudio La Selva).

HTML estático, sin build. Vercel la redespliega en cada push a `main`.

## Archivos

- `index.html` — la página entera: estilo, Tucu en Three.js, reglamento y formulario
- `logo-travesias.png` — logotipo en crema con transparencia

## Qué hay que editar

Todo lo configurable vive al inicio del primer `<script>`:

| Constante  | Para qué |
|------------|----------|
| `BACKEND`  | URL `/exec` del Apps Script. **Vacía = modo borrador**: el formulario valida pero no envía nada |
| `CIERRE`   | Fecha límite (29 de enero de 2027, 23:59 hora de Colombia) |
| `ENCUADRE` | Encuadre de Tucu, en unidades de cráneo |
| `MIRA_X/Y` | Signo del seguimiento de cabeza |

El reglamento y los campos del formulario son marcado directo, con los
pendientes señalados como tales.

## Subida de archivos

El participante **no necesita cuenta de Google**. El Apps Script corre como
`equipo@estudiolaselva.com`, abre una sesión de subida reanudable contra Drive
y devuelve una URL de un solo uso; el navegador sube los bytes directo a Google
sin pasar por el script, así que no hay techo de tamaño.

Dos cosas que rompen esto si se tocan:

1. Las peticiones al script van como `text/plain` a propósito — con
   `application/json` el navegador exige un preflight que Apps Script no responde.
2. Google amarra el permiso CORS de la sesión al `Origin` declarado **al crearla**.
   El dominio de esta página tiene que estar en la lista `ORIGENES` del script.
