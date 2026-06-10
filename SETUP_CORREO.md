# Configurar los reportes automáticos por correo

El dashboard envía dos correos automáticos con el resumen de desempeño (equipo + por
coordinadora + incidencias de la semana/mes):

- **Semanal** — cada lunes a las ~8:00 a.m. (hora del centro), con la **semana previa**.
- **Mensual** — el día 1 de cada mes a las ~8:00 a.m., con el **mes previo**.

Los correos se mandan con **Resend** y se envían a la lista de personas que tú definas.
Solo hay que configurarlo **una vez**. No requiere saber programar.

> ⚠️ **Importante:** el correo usa los datos del archivo `datos.xlsx` que esté en el
> repositorio al momento del envío. Sube el export de Atom **antes** de cada lunes para que
> el reporte semanal salga con datos actualizados.

---

## Paso 1 · Crear una cuenta en Resend (gratis)

1. Entra a **https://resend.com** y crea una cuenta (plan gratuito: 3,000 correos/mes).
2. En **Domains**, verifica un dominio de correo (ideal: el de la clínica), o usa el dominio
   de prueba que Resend ofrece para empezar.
3. En **API Keys → Create API Key**, copia la clave (empieza con `re_...`). Guárdala, solo se
   muestra una vez.

## Paso 2 · Guardar 3 datos secretos en GitHub

En el repositorio: **Settings → Secrets and variables → Actions → New repository secret**.
Crea estos tres (uno por uno):

| Nombre del secreto | Valor |
|---|---|
| `RESEND_API_KEY` | La clave `re_...` de Resend |
| `REPORT_FROM` | El remitente verificado, p. ej. `Auditoría FI <reportes@tudominio.com>` |
| `REPORT_TO` | Los destinatarios separados por coma, p. ej. `direccion@fi.com, supervisora@fi.com` |

## Paso 3 · ¡Listo!

Los correos saldrán solos en los horarios indicados.

### Probarlo ahora mismo (sin esperar al lunes)

1. Ve a la pestaña **Actions** del repositorio.
2. Elige **“Reportes PXI por correo”** en la lista de la izquierda.
3. Botón **Run workflow** → elige `week` o `month` → **Run workflow**.
4. En ~1 minuto te llega el correo de prueba.

---

## Cambiar destinatarios u horarios

- **Destinatarios:** edita el secreto `REPORT_TO` (Paso 2).
- **Horarios:** se definen en `.github/workflows/reports.yml` (líneas `cron`). Avísame y te
  los ajusto a la hora que prefieras.
