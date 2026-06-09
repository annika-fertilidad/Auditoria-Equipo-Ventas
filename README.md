# Índice de Experiencia del Paciente (PXI)

Dashboard de auditoría de conversaciones de ventas (WhatsApp / Atom) para Fertilidad Integral.
Calcula un **PXI por agente** a partir de 5 pilares auto-evaluados + 4 alertas de intervención
inmediata + una lista de muestreo manual para el supervisor.

## ¿Cómo funciona?

El panel carga automáticamente el archivo **`datos.xlsx`** o **`datos.csv`** del repositorio
(export nocturno de Atom, hoja `Historial`). Al abrir el enlace ves todo de inmediato — sin subir nada.

- **Scorecard PXI** — PXI por agente, los 5 pilares, SLA, mediana de respuesta y citas. Más gráficas.
- **Intervención inmediata** — cola de alertas R1–R4 con enlace directo a Atom, filtrable.
- **Muestreo manual** — checklist que el supervisor llena sobre ~3 chats por agente (se guarda en el navegador).

## Los 5 pilares (auto-evaluados)

| Pilar | Peso | Aplica cuando |
|---|---|---|
| P1 · Velocidad | 20% | Hay primer mensaje entrante + respuesta humana dentro de 2 días |
| P2 · Atención plena | 25% | Siempre (falla si el paciente quedó esperando respuesta) |
| P3 · Valor antes de precio | 15% | El agente menciona un precio |
| P4 · Lenguaje seguro | 25% | Siempre |
| P5 · Sensibilidad emocional | 15% | El paciente expresó carga emocional o frustración |

**PXI** = promedio ponderado de los pilares aplicables (re-normalizado). Agentes con <3
conversaciones quedan fuera del ranking.

## Alertas de intervención (R1–R4, no entran al PXI)

- **R1** Frustración no validada
- **R2** Información clínica sin confirmación del equipo médico
- **R3** Cotización de tratamiento de alto valor (FIV/PGT/ovodonación) sin involucrar supervisora
- **R4** Lead interesado >24h sin respuesta

## Muestreo manual (no se califica automáticamente)

El export es **solo texto** — precios, paquetes e info en imágenes/PDF no aparecen. Por eso estos
puntos los revisa el supervisor a mano: uso real del nombre, profundidad SPIN, protocolo de
objeciones, siguiente paso fechado, documentación en HubSpot.

## Cómo actualizar los datos (sin programar)

1. Exporta de Atom el archivo nocturno (`.xlsx`, hoja `Historial`).
2. Renómbralo a **`datos.xlsx`** (nombre exacto).
3. En GitHub: **Add file → Upload files**, arrastra el archivo, **Commit changes**.
4. En ~1 minuto el dashboard se actualiza.

> El botón **Cargar export** permite revisar un archivo local sin subirlo a GitHub.

## Columnas esperadas (hoja `Historial`)

`num_conversacion, cliente_csv, contacto, fecha_inicio_gestion, canal, agente, tipificacion,
es_venta, tipo, direccion, remitente, contenido, hora, url`

## Publicar en GitHub Pages

**Settings → Pages → Source → main / root**. Quedará en
`https://annika-fertilidad.github.io/Auditoria-Equipo-Ventas`.
