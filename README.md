# Auditoría Equipo de Ventas

Panel de control para auditorías semanales de chats de WhatsApp del equipo de ventas.

## ¿Cómo funciona?

El panel carga automáticamente el archivo **`datos.csv`** que vive en este repositorio.
Cuando abres el enlace del dashboard, ves de inmediato todo el histórico — sin tener
que subir nada.

- **Pestaña "Cumplimiento semanal"** — cumplimiento por agente, criterios más fallidos
  y evolución por semana.
- **Pestaña "Incidencias"** — qué incidencias ocurrieron, cuántas y en qué chats.
- **Filtro por semana y por agente** en la parte superior.

## Cómo actualizar los datos (sin saber programar)

1. Abre la **plantilla** desde el botón ⬇ Plantilla del dashboard y llénala con las
   auditorías nuevas (o agrégalas a tu archivo existente).
2. Guarda/exporta tu archivo como **`datos.csv`** (mismo nombre exacto).
3. En GitHub, entra a este repositorio y haz clic en **Add file → Upload files**.
4. Arrastra tu nuevo `datos.csv`, marca **"Commit changes"** y listo.
5. En ~1 minuto el dashboard mostrará los datos actualizados.

> El botón **⬆ Cargar archivo** sigue disponible para revisar un archivo en tu
> computadora sin subirlo a GitHub.

## Columnas del archivo `datos.csv`

| Columna | Valores |
|---|---|
| Fecha | Fecha de la auditoría (AAAA-MM-DD) |
| Agente | Nombre del agente |
| Chat_ID | Identificador del chat (ej. Chat_001) |
| C1_Nombre … C8_InfoMedica | `Sí` / `No` |
| C9_Emocion, C10_EVRAG | `Sí` / `No` / `N/A` (N/A cuando no hubo situación emocional) |
| I1_PrecioSinValor … I8_NoDocumentoHubSpot | `Sí` (ocurrió la incidencia) / `No` |

**Criterios (C):**
1. Usó el nombre del paciente desde el primer mensaje
2. Hizo al menos 2 preguntas abiertas (SPIN) antes de dar información
3. Presentó el valor antes del precio (Trío del Valor)
4. La conversación terminó con un siguiente paso concreto y fechado
5. Respondió dentro de los 2 minutos establecidos
6. Documentó correctamente en HubSpot
7. Aplicó el protocolo de objeciones (Reconocer→Validar→Informar→Invitar)
8. Evitó dar información médica sin confirmación del equipo clínico
9. Reconoció la emoción del paciente antes de pasar a información
10. Aplicó el protocolo EVRA+G en caso de frustración

**Incidencias (I):**
1. Presentó el precio sin presentar valor primero
2. Lead con interés real >24 hs sin respuesta
3. Usó la plantilla genérica sin personalización
4. Frustración del paciente no validada emocionalmente
5. Cotizó tratamiento de alto valor sin involucrar a la supervisora
6. Cerró el deal como perdido antes de los 2 intentos
7. Dio información clínica sin confirmación del equipo médico
8. No documentó la conversación en HubSpot

## Publicar en GitHub Pages

1. Ve a **Settings → Pages → Source → main branch / root**.
2. El panel quedará disponible en `https://annika-fertilidad.github.io/Auditoria-Equipo-Ventas`.
