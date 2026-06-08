# Auditoría Equipo de Ventas

Panel de control para auditorías diarias de chats de WhatsApp del equipo de ventas.

## ¿Cómo usar?

1. Abre `index.html` en tu navegador (o publícalo en GitHub Pages).
2. Descarga la **plantilla CSV** desde el botón en la esquina superior derecha.
3. Llena la plantilla con los resultados de cada auditoría:
   - **Fecha** — fecha de la auditoría (YYYY-MM-DD o DD/MM/YYYY)
   - **Agente** — nombre del agente evaluado
   - **Categorías** — puntaje de 0 a 100 por cada categoría del manual
   - **Puntaje Total** — promedio general (se calcula automáticamente si se omite)
4. Carga el archivo en el panel con el botón **Cargar archivo**.

## Columnas de la plantilla

| Columna | Descripción |
|---|---|
| Fecha | Fecha de la auditoría |
| Agente | Nombre del agente |
| Saludo y Presentación | Puntaje 0–100 |
| Manejo de Objeciones | Puntaje 0–100 |
| Lenguaje y Tono | Puntaje 0–100 |
| Seguimiento al Cliente | Puntaje 0–100 |
| Cierre de Venta | Puntaje 0–100 |
| Puntaje Total | Promedio general (opcional) |

> Puedes agregar o quitar columnas de categorías según tu manual de ventas.

## Funcionalidades

- Filtro por agente y rango de fechas
- KPIs: puntaje promedio, total auditorías, agentes evaluados, tasa de aprobación
- Gráfica de barras por agente
- Evolución de puntajes en el tiempo
- Radar de promedios por categoría
- Distribución de puntajes
- Tabla ordenable y exportable

## Publicar en GitHub Pages

1. Sube los archivos a un repositorio de GitHub.
2. Ve a **Settings → Pages → Source → main branch / root**.
3. El panel estará disponible en `https://tu-usuario.github.io/nombre-repo`.
