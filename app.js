/* ── Patient Experience Index (PXI) · Scoring engine + dashboard ── */

// ── Config ──────────────────────────────────────────────────────────────────
const WEIGHTS = { p1: 20, p2: 20, p3: 15, p4: 20, p5: 10, p6: 15 };

const BOTS = ['fi bot','atom','api','agendamiento - instagram','agendamiento - facebook'];

const TOK = {
  request: ['precio','costo','cuanto','como','cuando','donde','cual','puedo','pueden','podria',
    'informacion','info','agendar','cita','quiero','necesito','me interesa','estoy interesad',
    'disponib','horario','direccion','mandame','enviame','me puedes','una pregunta','una duda','sirve','aplica'],
  closer: ['gracias','ok','vale','perfecto','listo','igualmente','bendiciones','de nada','claro','excelente'],
  value: ['consulta','evaluacion','valoracion','diagnostico','plan','incluye','especialista','estudio',
    'paquete','programa','ultrasonido','antimulleriana','reserva ovarica','inseminacion','in vitro','fiv',
    'congelamiento','ovodonacion','espermatobioscopia','fragmentacion','tratamiento','hormona','laboratorio',
    'perfil','check','revision','pgt','dgp','prueba','embrion','criopreserv','biopsia','muestra','analisis',
    'seleccion','sesion','servicio','contiene','consiste','abarca','comprende','contempla'],
  clinical: ['tasa de exito','% de exito','porcentaje de exito','probabilidad de embarazo','vas a lograr',
    'vas a quedar embarazada','te garantizo','garantizamos el embarazo','eres buen candidat','eres buena candidat',
    'tu diagnostico es'],
  prohibited: ['relajate','no te estreses','todavia eres joven','al menos puedes','todo pasa por algo',
    'muchas personas pasan por esto','infertil','ciclo fallido','fallo el ciclo','embarazo geriatrico',
    'todo va a estar bien','todos los casos tienen solucion','es una inversion en su futuro',
    'es una inversion en tu futuro','cada mes cuenta','a su edad no puede esperar','a tu edad no puede esperar',
    'si no actua ahora','madre de alquiler','vientre de alquiler'],
  emotion: ['anos intentando','anos buscando','llevamos anos','perdida gestacional','aborto','perdi a mi bebe',
    'perdi el embarazo','desesper','no he podido embaraz','no hemos podido embaraz','no puedo quedar embaraz',
    'me siento triste','muy frustrad','estoy agotada','cansados de','mucha angustia'],
  frustration: ['muy molesta','muy molesto','estoy molesta','estoy molesto','enojad','indignad','pesimo servicio',
    'mal servicio','nadie me ha','llevo esperando','sigo esperando','no me han contestado','no me han respondido',
    'pesima atencion','inaceptable','decepcion'],
  validation: ['entiendo','comprendo','es valid','por supuesto','sin compromiso','te entiendo','lo siento',
    'siento mucho','lamento','una disculpa','que valiente','estamos contigo','estamos aqui','aqui estamos'],
  highValue: ['fiv','in vitro','pgt','dgp','ovodonacion'],
};

const PILLAR_LABELS = {
  p1: 'P1 · Velocidad', p2: 'P2 · Atención plena', p3: 'P3 · Valor antes de precio',
  p4: 'P4 · Lenguaje seguro', p5: 'P5 · Calidad de redacción', p6: 'P6 · Amabilidad y cortesía',
};

// Chat-speak / informal abbreviations that look unprofessional in writing.
const CHATSPEAK = ['xq','xk','pq','xfa','xfis','porfa','porfis','tmb','tb','bn','tqm','xd','salu2',
  'finde','ntp','nps','q','k','d','x','pa','pq','grax','graxias','holaa','siii','noo','aki','ke','komo'];
// Courtesy / friendliness markers (normalized, accent-free).
const GREET = ['hola','buen dia','buenos dias','buenas tardes','buenas noches','que tal','bienvenid',
  'que gusto saludar','un gusto saludar'];
const POLITE = ['por favor','porfavor','gracias','con gusto','con mucho gusto','claro que si','encantad',
  'quedo atent','estoy para ayudar','estoy para servir','con todo gusto','sera un placer','que tenga',
  'excelente dia','feliz dia','feliz tarde','un gusto','para servirte','no dude','cualquier duda',
  'quedo a la orden','a la orden','quedo pendiente','con todo gusto te','sera un placer atender','permiteme',
  'con gusto te','te comparto','te ayudo','te apoyo','te puedo ayudar','quedo al pendiente','estamos al pendiente',
  'estamos para','no te preocupes','no hay problema','con todo el gusto','sera un gusto','con cariño','un abrazo',
  'que tengas','saludos cordiales','muchas gracias','mil gracias','claro que','por nada','estamos en contacto'];

const MONTHS = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto',
  'septiembre','octubre','noviembre','diciembre'];

// Nombres que aparecen en la columna `agente` pero NO son coordinadoras de ventas
// (cuentas de sistema / contactos / personal ajeno al equipo). Se excluyen de
// todas las vistas. Comparación normalizada (sin acentos, minúsculas, sin dobles
// espacios) para que no importe cómo venga escrito en el export.
const NON_AGENTS = ['hola fertilidad', 'mariana sanchez'];
const normAgent = s => String(s || '')
  .normalize('NFD').replace(/[̀-ͯ]/g, '')
  .replace(/\s+/g, ' ').trim().toLowerCase();
const isNonAgent = name => NON_AGENTS.includes(normAgent(name));

// ── State ──────────────────────────────────────────────────────────────────
let conversations = [];      // full accumulated history
let weekConversations = [];  // current-week slice (scorecard / flags / sampling)
let agentStats = [];
let clinic = {};
let latestDayTs = null;      // most recent dated day
let currentWeekStart = null; // Monday of the most recent week
let currentMonthStart = null;// 1st of the most recent month
let activeTab = 'scorecard';
let flagFilterAgent = '', flagFilterType = '';

// ── Helpers ─────────────────────────────────────────────────────────────────
const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const hasAny = (text, tokens) => tokens.some(t => text.includes(t));
const isBot = remitente => BOTS.includes(norm(remitente).trim());

function parseHora(h) {
  // "DD/MM/YYYY a las HH:MM am/pm"
  const m = String(h || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4}).*?(\d{1,2}):(\d{2})\s*(am|pm)?/i);
  if (!m) return null;
  let [_, d, mo, y, hh, mm, ap] = m;
  hh = +hh; mm = +mm;
  if (ap) { ap = ap.toLowerCase(); if (ap === 'pm' && hh < 12) hh += 12; if (ap === 'am' && hh === 12) hh = 0; }
  return new Date(+y, +mo - 1, +d, hh, mm).getTime();
}

// ── File handling ──────────────────────────────────────────────────────────
function handleFile(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false });
      ingestWorkbook(wb, file.name);
    } catch (err) { showToast('Error al leer el archivo: ' + err.message); }
  };
  reader.readAsArrayBuffer(file);
  evt.target.value = '';
}

function ingestWorkbook(wb, filename) {
  const sheetName = wb.SheetNames.find(n => norm(n) === 'historial') || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (!rows.length) { showToast('El archivo está vacío.'); return; }
  buildConversations(rows);
  if (!conversations.length) { showToast('No se encontraron conversaciones de ventas válidas.'); return; }
  aggregate();
  renderAll();
  showToast(`✓ ${conversations.length} conversaciones auditadas (${filename})`);
  document.getElementById('uploadZone').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('tabBar').style.display = 'flex';
}

// ── Build conversations ─────────────────────────────────────────────────────
function col(row, names) {
  for (const k of Object.keys(row)) if (names.includes(norm(k).trim())) return row[k];
  return '';
}

function buildConversations(rows) {
  const groups = {};
  rows.forEach(r => {
    // OJO: `num_conversacion` NO es único — Atom reutiliza esos números entre
    // conversaciones distintas, así que agrupar por él mezcla pacientes
    // diferentes y genera huecos falsos (p. ej. "sin respuesta >2h"). El
    // identificador real de la conversación es la URL de Atom. Agrupamos por
    // URL; si faltara, caemos al num_conversacion.
    const num = String(col(r, ['num_conversacion','num conversacion','conversacion','id'])).trim();
    const url = String(col(r, ['url'])).trim();
    const key = url || num;
    if (!key) return;
    (groups[key] ||= []).push({
      tipo: norm(col(r, ['tipo'])),
      direccion: norm(col(r, ['direccion'])),
      remitente: String(col(r, ['remitente'])).trim(),
      contenido: String(col(r, ['contenido'])),
      hora: parseHora(col(r, ['hora'])),
      agente: String(col(r, ['agente'])).trim(),
      tipificacion: String(col(r, ['tipificacion'])).trim(),
      es_venta: norm(col(r, ['es_venta','es venta'])),
      cliente: String(col(r, ['remitente'])).trim(),
      num: num,
      url: url,
    });
  });

  conversations = [];
  Object.entries(groups).forEach(([key, msgs]) => {
    const num = msgs.find(m => m.num)?.num || key;
    // messages only carry text
    let chat = msgs.filter(m => m.tipo === 'mensaje');
    chat.sort((a, b) => (a.hora || 0) - (b.hora || 0));
    const inbound  = chat.filter(m => m.direccion === 'entrante');
    const humanOut = chat.filter(m => m.direccion === 'saliente' && !isBot(m.remitente));
    if (!inbound.length || !humanOut.length) return; // not sales-relevant

    // Coordinator = the assigned owner in the `agente` column (per spec).
    // Whitespace is normalized so "Michelle  Hernandez" and "Michelle Hernandez" don't split.
    // Fallback: the human (non-bot) person who sent the outbound replies. Patient names live on
    // `entrante` messages and are never used here.
    const cleanName = s => String(s || '').replace(/\s+/g, ' ').trim();
    const agenteCol = cleanName(msgs.find(m => cleanName(m.agente))?.agente);
    const outCounts = {};
    humanOut.forEach(m => { const n = cleanName(m.remitente); if (n) outCounts[n] = (outCounts[n] || 0) + 1; });
    const topOut = Object.entries(outCounts).sort((a, b) => b[1] - a[1])[0]?.[0];
    const agente = agenteCol || topOut || 'Sin asignar';
    if (isNonAgent(agente)) return; // cuenta de sistema / no es coordinadora
    const url = msgs.find(m => m.url)?.url || '';
    const tipificacion = msgs.find(m => m.tipificacion)?.tipificacion || '';
    const es_venta = msgs.find(m => m.es_venta)?.es_venta || '';

    // Atom assignment/handoff events ("ATOM/X ha (re)asignado esta conversación … a <agente>").
    // tipo='evento' (filtered out of `chat`), so we collect them here. These mark the moment
    // the human agent actually receives the chat — the correct start for the velocity (P1) clock.
    const assignTs = msgs
      .filter(m => m.tipo === 'evento' && /asignad[oa]/i.test(m.contenido) && !/al bot/i.test(m.contenido) && m.hora)
      .map(m => m.hora);

    conversations.push(scoreConversation({ num, chat, inbound, humanOut, agente, url, tipificacion, es_venta, assignTs }));
  });
}

// Working hours = 07:00–24:00 daily; off-hours = 00:00–07:00. The SLA clock
// only runs during working hours (it pauses overnight).
const WORK_START_H = 7;     // 7 a.m. (working hours run 7 a.m.–midnight, full coverage)
function workingMinutesBetween(t0, t1) {
  if (!t0 || !t1 || t1 <= t0) return 0;
  let total = 0;
  let cur = new Date(t0);
  while (cur.getTime() < t1) {
    const dayStart = new Date(cur); dayStart.setHours(WORK_START_H, 0, 0, 0);
    const nextMidnight = new Date(cur); nextMidnight.setHours(0, 0, 0, 0); nextMidnight.setDate(nextMidnight.getDate() + 1);
    const winStart = Math.max(cur.getTime(), dayStart.getTime());
    const winEnd = Math.min(t1, nextMidnight.getTime());
    if (winEnd > winStart) total += winEnd - winStart;
    cur = nextMidnight;
  }
  return Math.round(total / 60000);
}

// ── Score a single conversation ──────────────────────────────────────────────
function needsReply(text) {
  let t = norm(text).trim();
  if (!t) return false;
  if (t.includes('?') || String(text).includes('¿')) return true;
  // "en cuanto" / "en cuanto a" = "tan pronto como": NO es solicitud de precio.
  // Sin este guard, "En cuanto realice la transferencia…" matcheaba "cuanto" y
  // se marcaba como pregunta pendiente (falso positivo en chats ya cerrados).
  t = t.replace(/\ben cuanto\b/g, ' ');
  // Cierre/cortesía sin pregunta ("igualmente", "muchas gracias") no requiere respuesta.
  if (hasAny(t, TOK.closer) && !hasAny(t, TOK.request)) return false;
  if (hasAny(t, TOK.request)) return true;
  return false;
}

function scoreConversation(c) {
  const { chat, inbound, humanOut, assignTs } = c;
  const agentText = norm(humanOut.map(m => m.contenido).join('  ||  '));
  const patientText = norm(inbound.map(m => m.contenido).join('  ||  '));
  const pillars = {};

  // ── P1 Speed ── (working hours 7am–midnight; clock pauses overnight)
  // Two-tier SLA: first reply must be < 2 min; every reply afterwards within 15 min.
  // El reloj de velocidad arranca cuando la asesora RECIBE el chat (handoff de Atom),
  // no en el primer mensaje del lead: el bot atiende la calificación inicial al instante
  // y el tiempo de enrutamiento no es responsabilidad de la asesora. Si no hay evento de
  // asignación, caemos al primer mensaje entrante.
  const firstIn = inbound[0]?.hora;
  const firstReply = humanOut.find(m => m.hora && firstIn && m.hora >= firstIn)?.hora;
  let slaStart = firstIn;
  if (firstReply != null && firstIn != null && Array.isArray(assignTs) && assignTs.length) {
    const handoffs = assignTs.filter(t => t != null && t <= firstReply);
    if (handoffs.length) slaStart = Math.max(firstIn, ...handoffs);
  }
  let firstRespMin = (slaStart && firstReply) ? workingMinutesBetween(slaStart, firstReply) : null;

  // Subsequent reply gaps (after the first reply), in working minutes.
  let maxSubRespMin = null;
  let pSince = null, seenFirstReply = false;
  for (const m of chat) {
    if (!m.hora) continue;
    if (m.direccion === 'entrante') {
      if (pSince === null && needsReply(m.contenido)) pSince = m.hora;
    } else if (m.direccion === 'saliente') {
      if (!seenFirstReply) { seenFirstReply = true; pSince = null; continue; }
      if (pSince !== null) {
        const wm = workingMinutesBetween(pSince, m.hora);
        maxSubRespMin = maxSubRespMin === null ? wm : Math.max(maxSubRespMin, wm);
        pSince = null;
      }
    }
  }

  const firstOK = firstRespMin !== null && firstRespMin <= 2;
  const subOK = maxSubRespMin === null || maxSubRespMin <= 15;
  if (firstRespMin !== null) {
    pillars.p1 = { applies: true, pass: firstOK && subOK };
  } else {
    pillars.p1 = { applies: false, pass: null };
  }
  c.firstRespMin = firstRespMin;          // working minutes to first reply
  c.maxSubRespMin = maxSubRespMin;        // worst subsequent reply gap (working min)
  c.repliedUnder2 = firstRespMin !== null && firstRespMin <= 2;

  // ── P2 Full Attention ── (always applies; FAIL if patient left hanging)
  const last = chat[chat.length - 1];
  const dropped = last && last.direccion === 'entrante' && needsReply(last.contenido);
  pillars.p2 = { applies: true, pass: !dropped };

  // ── P3 Value before Price ──
  const priceRe = /\$\s?\d|\b\d{3,}\b\s*(pesos|mxn|mil)|cuesta|tiene un costo|el costo es|el precio es|son \$/i;
  let p3Applies = false, p3Pass = null, gavePrice = false;
  for (let i = 0; i < humanOut.length; i++) {
    if (priceRe.test(norm(humanOut[i].contenido)) || /\$\s?\d/.test(humanOut[i].contenido)) {
      gavePrice = true; p3Applies = true;
      const before = norm(humanOut.slice(0, i + 1).map(m => m.contenido).join(' '));
      const after  = norm(humanOut.slice(i + 1, i + 3).map(m => m.contenido).join(' '));
      const ctx = hasAny(before, TOK.value) || hasAny(after, TOK.value);
      if (ctx) { p3Pass = true; }
      else if (p3Pass === null) { p3Pass = false; }
    }
  }
  pillars.p3 = { applies: p3Applies, pass: p3Applies ? p3Pass : null };
  c.gavePrice = gavePrice;

  // ── P4 Safe Language ── (always applies)
  let clinicalHit = false;
  for (const claim of TOK.clinical) {
    let idx = agentText.indexOf(claim);
    while (idx !== -1) {
      const pre = agentText.slice(Math.max(0, idx - 12), idx);
      if (!/\b(no|sin)\s$/.test(pre)) { clinicalHit = true; break; }
      idx = agentText.indexOf(claim, idx + 1);
    }
    if (clinicalHit) break;
  }
  const prohibitedHit = hasAny(agentText, TOK.prohibited);
  pillars.p4 = { applies: true, pass: !(clinicalHit || prohibitedHit) };
  c.clinicalHit = clinicalHit;
  c.prohibitedHit = prohibitedHit;

  // (Emotional sensitivity now lives on as the R1 alert, below.)
  const hasFrustration = hasAny(patientText, TOK.frustration);
  const validated = hasAny(agentText, TOK.validation);
  c.hasFrustration = hasFrustration;
  c.validated = validated;

  // ── P5 Writing Quality ── (style heuristic on the agent's own messages)
  // Judge only "substantial" messages (≥3 words) so short "¡Hola! 😊" replies aren't penalized.
  const substantial = humanOut.filter(m => {
    const raw = (m.contenido || '').trim();
    return raw.replace(/\s+/g, ' ').split(' ').length >= 3 && raw.replace(/[^a-záéíóúñ]/gi, '').length >= 12;
  });
  const writingIssues = [];
  let badMsgs = 0;
  for (const m of substantial) {
    const raw = (m.contenido || '').trim();
    const letters = raw.replace(/[^a-záéíóúñ]/gi, '');
    const issues = [];
    // 1) Sentence doesn't start with a capital letter
    const firstLetter = raw.match(/[a-záéíóúñ]/i);
    if (firstLetter && firstLetter[0] === firstLetter[0].toLowerCase()) issues.push('mayúscula inicial');
    // 2) Chat-speak abbreviations (whole-word match, punctuation stripped)
    const tokN = ' ' + norm(raw).replace(/[^a-z0-9ñ ]/gi, ' ').replace(/\s+/g, ' ').trim() + ' ';
    if (CHATSPEAK.some(w => tokN.includes(' ' + w + ' '))) issues.push('abreviaturas informales');
    // 3) ALL-CAPS shouting
    if (letters.length >= 8) {
      const ups = (raw.match(/[A-ZÁÉÍÓÚÑ]/g) || []).length;
      if (ups / letters.length > 0.8) issues.push('mayúsculas sostenidas');
    }
    // 4) Question mark without the opening ¿
    if (raw.includes('?') && !raw.includes('¿')) issues.push('falta ¿ de apertura');
    if (issues.length) { badMsgs++; issues.forEach(i => { if (!writingIssues.includes(i)) writingIssues.push(i); }); }
  }
  if (substantial.length >= 1) {
    pillars.p5 = { applies: true, pass: (badMsgs / substantial.length) <= 0.2 };
  } else {
    pillars.p5 = { applies: false, pass: null };
  }
  c.writingIssues = writingIssues;
  c.writingBadRate = substantial.length ? Math.round(badMsgs / substantial.length * 100) : null;

  // ── P6 Friendliness & Courtesy ── (always applies)
  const greeted = hasAny(agentText, GREET);
  const polite = hasAny(agentText, POLITE) || validated;
  pillars.p6 = { applies: true, pass: greeted && polite };
  c.greeted = greeted;
  c.polite = polite;

  // ── Response-time analysis (only counts time the LEAD is waiting on the agent) ──
  // Una pregunta del lead arranca el reloj; la respuesta del agente lo detiene.
  // Reglas (afinadas para no marcar falsos >2h):
  //  · Solo cuentan los retrasos del MISMO día. Un hueco que cruza al día
  //    siguiente = lead que se enfrió y se re-enganchó después, NO un
  //    incumplimiento de SLA (>2h).
  //  · Los mensajes que llegan en horario nocturno (12am–7am) entran a la cola
  //    matutina: deben atenderse en la 1ª hora del turno (antes de las 8am);
  //    si se responden después, sí es tardío.
  //  · Una pregunta en horario laboral que queda sin responder al final del
  //    snapshot NO se marca como >2h (pudo responderse tras el corte del export);
  //    eso ya lo captura P2 (atención plena).
  const sameDay = (a, b) => new Date(a).toDateString() === new Date(b).toDateString();
  let maxWaitWorking = 0, pendingSince = null;
  let agentSlow = false, morningQueue = false;
  c.pendingHour = null;
  for (const m of chat) {
    if (!m.hora) continue;
    if (m.direccion === 'entrante') {
      if (pendingSince === null && needsReply(m.contenido)) pendingSince = m.hora;
    } else if (m.direccion === 'saliente' && pendingSince !== null) {
      const reply = m.hora;
      const h = new Date(pendingSince).getHours();
      if (h < WORK_START_H) {
        // Cola matutina: debió responderse antes de las 8am.
        maxWaitWorking = Math.max(maxWaitWorking, workingMinutesBetween(pendingSince, reply));
        const deadline = new Date(pendingSince); deadline.setHours(WORK_START_H + 1, 0, 0, 0);
        if (reply > deadline.getTime()) agentSlow = true;
      } else if (sameDay(pendingSince, reply)) {
        const wm = workingMinutesBetween(pendingSince, reply);
        maxWaitWorking = Math.max(maxWaitWorking, wm);
        if (wm > 120) agentSlow = true;
      } // else: respondió otro día → lead frío / re-enganche, sin incumplimiento.
      pendingSince = null;
    }
  }
  if (pendingSince !== null) {            // hilo termina con el lead esperando
    const h = new Date(pendingSince).getHours();
    c.pendingHour = h;
    if (h < WORK_START_H) morningQueue = true;  // llegó de noche → cola matutina (R5)
    // En horario laboral sin responder al cierre del snapshot: lo captura P2,
    // no lo duplicamos como R4/R5.
  }
  c.maxWaitWorking = maxWaitWorking;
  c.agentSlow = agentSlow;
  c.morningQueue = morningQueue;

  // ── Red flags ──
  c.flags = {
    r1: hasFrustration && !validated,
    r2: clinicalHit,
    r3: gavePrice && hasAny(norm(chat.map(m => m.contenido).join(' ')), TOK.highValue),
    r4: agentSlow,
    r5: morningQueue,
  };

  c.pillars = pillars;

  // Day + week buckets from the first dated message (a chat belongs to the day
  // the lead first wrote — stable even if the conversation continues later).
  const ts = chat.find(m => m.hora)?.hora || null;
  if (ts) {
    const dd = new Date(ts); dd.setHours(0, 0, 0, 0);
    c.dayTs = dd.getTime();
    c.dayLabel = `${String(dd.getDate()).padStart(2,'0')}/${String(dd.getMonth()+1).padStart(2,'0')}`;
    const d = new Date(ts);
    const day = (d.getDay() + 6) % 7;          // 0 = Monday
    d.setHours(0, 0, 0, 0); d.setDate(d.getDate() - day);
    c.weekStart = d.getTime();
    c.weekLabel = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
    const md = new Date(ts);
    const mo = new Date(md.getFullYear(), md.getMonth(), 1);
    c.monthStart = mo.getTime();
    c.monthLabel = MONTHS[md.getMonth()] + ' ' + md.getFullYear();
  } else { c.dayTs = null; c.dayLabel = 'Sin fecha'; c.weekStart = null; c.weekLabel = 'Sin fecha'; c.monthStart = null; c.monthLabel = 'Sin fecha'; }

  // PXI for this conversation
  let wSum = 0, sSum = 0;
  for (const k of Object.keys(WEIGHTS)) {
    if (pillars[k].applies) { wSum += WEIGHTS[k]; sSum += WEIGHTS[k] * (pillars[k].pass ? 1 : 0); }
  }
  c.pxi = wSum ? Math.round(sSum / wSum * 100) : null;
  return c;
}

// ── Aggregate per agent + clinic ────────────────────────────────────────────
function median(arr) {
  if (!arr.length) return null;
  const s = [...arr].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : Math.round((s[m - 1] + s[m]) / 2);
}

function aggregate() {
  // Current period = the week (and latest day) of the most recent dated chat.
  // Scoping the scorecard to the current week means each daily upload accumulates
  // toward the week, and an agent on a rest day still shows their weekly score.
  const dated = conversations.filter(c => c.dayTs != null);
  // "Semana actual": la semana más reciente con volumen real. Si la semana en
  // curso apenas empezó (p. ej. lunes con un solo chat cargado), mostrarla
  // dejaría el tablero casi vacío (una sola agente). Por eso elegimos la
  // semana más reciente con al menos MIN_CURRENT_WEEK conversaciones, y solo
  // si ninguna alcanza el umbral usamos la última semana disponible.
  const MIN_CURRENT_WEEK = 10;
  const weekCounts = {};
  dated.forEach(c => { weekCounts[c.weekStart] = (weekCounts[c.weekStart] || 0) + 1; });
  const weekStarts = Object.keys(weekCounts).map(Number).sort((a, b) => b - a);
  currentWeekStart = weekStarts.find(w => weekCounts[w] >= MIN_CURRENT_WEEK)
    ?? (weekStarts.length ? weekStarts[0] : null);
  // El día "actual" = el día más reciente DENTRO de la semana elegida, para que
  // la tarjeta "PXI de hoy" sea coherente con la semana mostrada.
  const inWeek = currentWeekStart != null
    ? dated.filter(c => c.weekStart === currentWeekStart) : dated;
  latestDayTs = inWeek.length ? Math.max(...inWeek.map(c => c.dayTs)) : null;
  // El mes actual se elige igual: el mes más reciente con datos (los meses
  // acumulan mucho más, así que basta con el más reciente).
  currentMonthStart = dated.length ? Math.max(...dated.map(c => c.monthStart)) : null;
  weekConversations = currentWeekStart != null
    ? conversations.filter(c => c.weekStart === currentWeekStart)
    : conversations.slice();
  const monthConversations = currentMonthStart != null
    ? conversations.filter(c => c.monthStart === currentMonthStart)
    : conversations.slice();
  const monthByAgent = {};
  monthConversations.forEach(c => (monthByAgent[c.agente] ||= []).push(c));

  const byAgent = {};
  weekConversations.forEach(c => (byAgent[c.agente] ||= []).push(c));

  agentStats = Object.entries(byAgent).map(([agent, convs]) => {
    const pillarPct = {};
    for (const k of Object.keys(WEIGHTS)) {
      const appl = convs.filter(c => c.pillars[k].applies);
      const pass = appl.filter(c => c.pillars[k].pass);
      pillarPct[k] = appl.length ? Math.round(pass.length / appl.length * 100) : null;
    }
    // Weekly PXI = weighted avg of available pillarPct, renormalized
    let wSum = 0, sSum = 0;
    for (const k of Object.keys(WEIGHTS)) {
      if (pillarPct[k] !== null) { wSum += WEIGHTS[k]; sSum += WEIGHTS[k] * pillarPct[k]; }
    }
    const pxi = wSum ? Math.round(sSum / wSum) : null;

    // Daily PXI = today's slice only; null (= rest day) if no chats that day.
    const dayConvs = convs.filter(c => c.dayTs === latestDayTs);
    const dailyPxi = pxiOf(dayConvs);
    const restDay = dayConvs.length === 0;

    // Monthly PXI = the agent's whole current-month slice (full history, not the week).
    const monthConvs = monthByAgent[agent] || [];
    const monthlyPxi = pxiOf(monthConvs);
    const monthCount = monthConvs.length;

    const respTimes = convs.map(c => c.firstRespMin).filter(v => v !== null);
    const slaAppl = convs.filter(c => c.pillars.p1.applies);
    const slaPass = slaAppl.filter(c => c.pillars.p1.pass);
    const flags = convs.reduce((n, c) => n + Object.values(c.flags).filter(Boolean).length, 0);

    return {
      agent, count: convs.length, pillarPct, pxi,
      dailyPxi, restDay, dayCount: dayConvs.length,
      monthlyPxi, monthCount,
      medianResp: median(respTimes),
      slaPct: slaAppl.length ? Math.round(slaPass.length / slaAppl.length * 100) : null,
      under2Pct: respTimes.length ? Math.round(respTimes.filter(v => v <= 2).length / respTimes.length * 100) : null,
      appts: convs.filter(c => c.es_venta === 'si').length,
      flags,
    };
  }).sort((a, b) => (b.pxi ?? -1) - (a.pxi ?? -1));

  const ranked = agentStats.filter(a => a.count >= 3);
  const allResp = weekConversations.map(c => c.firstRespMin).filter(v => v !== null);
  const slaAppl = weekConversations.filter(c => c.pillars.p1.applies);
  const slaPass = slaAppl.filter(c => c.pillars.p1.pass);
  const dayConvs = weekConversations.filter(c => c.dayTs === latestDayTs);
  clinic = {
    convs: weekConversations.length,
    dayConvs: dayConvs.length,
    pxi: ranked.length ? Math.round(ranked.reduce((s, a) => s + a.pxi * a.count, 0) / ranked.reduce((s, a) => s + a.count, 0)) : null,
    dayPxi: pxiOf(dayConvs),
    medianResp: median(allResp),
    slaPct: slaAppl.length ? Math.round(slaPass.length / slaAppl.length * 100) : null,
    appts: weekConversations.filter(c => c.es_venta === 'si').length,
    flags: weekConversations.reduce((n, c) => n + Object.values(c.flags).filter(Boolean).length, 0),
    weekLabel: currentWeekStart != null ? (conversations.find(c => c.weekStart === currentWeekStart)?.weekLabel || null) : null,
    dayLabel: latestDayTs != null ? (conversations.find(c => c.dayTs === latestDayTs)?.dayLabel || null) : null,
    monthPxi: pxiOf(monthConversations),
    monthConvs: monthConversations.length,
    monthLabel: currentMonthStart != null ? (conversations.find(c => c.monthStart === currentMonthStart)?.monthLabel || null) : null,
  };
}

// ── Tabs ────────────────────────────────────────────────────────────────────
function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  ['scorecard','evolution','flags','sampling'].forEach(p =>
    document.getElementById('panel-' + p).style.display = p === tab ? '' : 'none');
  if (tab === 'evolution') renderEvolution();
}

// ── Render ──────────────────────────────────────────────────────────────────
let charts = {};
function destroyChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }

function renderAll() {
  renderKPIs();
  renderScorecards();
  renderCharts();
  renderFlags();
  renderSampling();
}

function pxiClass(v) { return v == null ? '' : v >= 80 ? 'success' : v >= 60 ? 'warning' : 'danger'; }

function renderKPIs() {
  document.getElementById('kpiGrid').innerHTML = `
    ${kpi('PXI del mes' + (clinic.monthLabel ? ' · ' + clinic.monthLabel : ''), (clinic.monthPxi ?? '—'), '/100', pxiClass(clinic.monthPxi))}
    ${kpi('PXI de la clínica · semana' + (clinic.weekLabel ? ' ' + clinic.weekLabel : ''), (clinic.pxi ?? '—'), '/100', pxiClass(clinic.pxi))}
    ${kpi('PXI de hoy' + (clinic.dayLabel ? ' · ' + clinic.dayLabel : ''), (clinic.dayPxi ?? '—'), '/100', pxiClass(clinic.dayPxi))}
    ${kpi('Conversaciones · mes', clinic.monthConvs, clinic.convs != null ? ` · ${clinic.convs} esta semana` : '')}
    ${kpi('Cumple SLA velocidad', clinic.slaPct != null ? clinic.slaPct + '%' : '—', '', pxiClass(clinic.slaPct))}
    ${kpi('Citas agendadas (semana)', clinic.appts, '')}
    ${kpi('Alertas abiertas', clinic.flags, '', clinic.flags > 0 ? 'danger' : 'success')}
  `;
}
function kpi(label, value, suffix, cls = '') {
  return `<div class="kpi-card ${cls}">
    <div class="kpi-label">${label}</div>
    <div class="kpi-value">${value}<span class="kpi-suffix">${suffix}</span></div>
  </div>`;
}

function renderScorecards() {
  const ranked = agentStats.filter(a => a.count >= 3);
  const low = agentStats.filter(a => a.count < 3);

  const cardHtml = a => {
    const cls = pxiClass(a.pxi);
    const pillars = Object.keys(WEIGHTS).map(k => {
      const v = a.pillarPct[k];
      const pc = v == null ? 'na' : v >= 80 ? 'success' : v >= 60 ? 'warning' : 'danger';
      return `<div class="pillar-cell ${pc}">
        <div class="pillar-name">${PILLAR_LABELS[k].split(' · ')[0]}</div>
        <div class="pillar-pct">${v == null ? 'N/A' : v + '%'}</div>
      </div>`;
    }).join('');
    const dcls = a.restDay ? 'na' : pxiClass(a.dailyPxi);
    const dayChip = a.restDay
      ? `<span class="sc-day na">Descanso hoy</span>`
      : `<span class="sc-day ${dcls}">Hoy: <strong>${a.dailyPxi ?? '—'}</strong> · ${a.dayCount} conv.</span>`;
    const mcls = a.monthlyPxi == null ? 'na' : pxiClass(a.monthlyPxi);
    const monthChip = `<span class="sc-day ${mcls}">Mes: <strong>${a.monthlyPxi ?? '—'}</strong> · ${a.monthCount} conv.</span>`;
    return `<div class="fi-card scorecard">
      <div class="sc-head">
        <div>
          <div class="sc-agent">${a.agent}</div>
          <div class="sc-sub">${a.count} conv. esta semana · ${a.appts} cita${a.appts !== 1 ? 's' : ''}</div>
          <div class="sc-dayrow">${dayChip} ${monthChip}</div>
        </div>
        <div class="sc-pxi ${cls}">${a.pxi ?? '—'}<span>PXI semana</span></div>
      </div>
      <div class="pillar-grid">${pillars}</div>
      <div class="sc-foot">
        <span>SLA vel.: <strong>${a.slaPct ?? '—'}%</strong></span>
        <span>Mediana: <strong>${a.medianResp != null ? a.medianResp + 'm' : '—'}</strong></span>
        <span class="${a.flags > 0 ? 'flag-warn' : ''}">Alertas: <strong>${a.flags}</strong></span>
      </div>
    </div>`;
  };

  document.getElementById('scorecardGrid').innerHTML = ranked.map(cardHtml).join('');
  const lowWrap = document.getElementById('lowVolume');
  if (low.length) {
    lowWrap.style.display = '';
    lowWrap.innerHTML = `<div class="section-note">Agentes con &lt;3 conversaciones (fuera del ranking)</div>
      <div class="scorecard-grid">${low.map(cardHtml).join('')}</div>`;
  } else lowWrap.style.display = 'none';
}

function renderCharts() {
  const ranked = agentStats.filter(a => a.count >= 3);
  const labels = ranked.map(a => a.agent);
  const SAGE = '#738D84', FOREST = '#20281B';
  const PIL_COLORS = ['#4A7B9D','#738D84','#C0574A','#6B9E6E','#8FA89F'];

  destroyChart('chartPxi');
  charts.chartPxi = new Chart(document.getElementById('chartPxi'), {
    type: 'bar',
    data: { labels, datasets: [{ label: 'PXI', data: ranked.map(a => a.pxi),
      backgroundColor: SAGE, borderRadius: 6 }] },
    options: baseOpts(100, '%'),
  });

  destroyChart('chartPillars');
  charts.chartPillars = new Chart(document.getElementById('chartPillars'), {
    type: 'bar',
    data: { labels, datasets: Object.keys(WEIGHTS).map((k, i) => ({
      label: PILLAR_LABELS[k].split(' · ')[1],
      data: ranked.map(a => a.pillarPct[k]),
      backgroundColor: PIL_COLORS[i], borderRadius: 4,
    })) },
    options: { ...baseOpts(100, '%'), plugins: { legend: { position: 'bottom', labels: { font: { size: 10 }, boxWidth: 10 } } } },
  });

  destroyChart('chartSla');
  charts.chartSla = new Chart(document.getElementById('chartSla'), {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Cumple SLA velocidad', data: ranked.map(a => a.slaPct),
      backgroundColor: '#6B9E6E', borderRadius: 6 }] },
    options: baseOpts(100, '%'),
  });
}

function baseOpts(max, suffix) {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      y: { min: 0, max, ticks: { callback: v => v + (suffix || ''), font: { size: 11 }, color: '#3D4E36' }, grid: { color: '#E8E7D8' } },
      x: { ticks: { font: { size: 11 }, color: '#3D4E36' }, grid: { display: false } },
    },
  };
}

// ── Weekly evolution ─────────────────────────────────────────────────────────
function pxiOf(convs) {
  let wSum = 0, sSum = 0;
  for (const k of Object.keys(WEIGHTS)) {
    const appl = convs.filter(c => c.pillars[k].applies);
    if (!appl.length) continue;
    const pct = appl.filter(c => c.pillars[k].pass).length / appl.length * 100;
    wSum += WEIGHTS[k]; sSum += WEIGHTS[k] * pct;
  }
  return wSum ? Math.round(sSum / wSum) : null;
}

function trendStatus(delta) {
  if (delta === null) return { label: 'Sin comparación', cls: 'na', arrow: '·' };
  if (delta >= 3)  return { label: 'Mejorando', cls: 'up',   arrow: '▲' };
  if (delta <= -3) return { label: 'Bajando',   cls: 'down', arrow: '▼' };
  return { label: 'Estable', cls: 'flat', arrow: '▬' };
}

// Etiqueta de semana CON año (DD/MM/AA) para evitar que semanas de meses/años
// distintos se vean revueltas en el eje.
function weekLabelYr(weekStart) {
  const d = new Date(weekStart);
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getFullYear()).slice(2)}`;
}

// Mínimo de conversaciones para que una semana cuente como "real". Las semanas
// con 1–2 conversaciones suelen ser leads antiguos que reescribieron (su primer
// mensaje cae en una semana vieja) y ensucian la gráfica.
const MIN_WEEK_CONVS = 3;

function renderEvolution() {
  // weeks sorted by date; only dated conversations
  const dated = conversations.filter(c => c.weekStart !== null);
  const weekCount = {};
  dated.forEach(c => { weekCount[c.weekStart] = (weekCount[c.weekStart] || 0) + 1; });
  // Solo semanas con datos reales, y como máximo las 12 más recientes.
  const weekKeys = Object.keys(weekCount).map(Number)
    .filter(k => weekCount[k] >= MIN_WEEK_CONVS)
    .sort((a, b) => a - b)
    .slice(-12);
  const weekLabels = weekKeys.map(k => 'Sem ' + weekLabelYr(k));
  const weekKeySet = new Set(weekKeys);
  const datedRecent = dated.filter(c => weekKeySet.has(c.weekStart));

  // Derive agents from the visible window (those active in the last 12 weeks, >=3 convs).
  const allByAgent = {};
  datedRecent.forEach(c => (allByAgent[c.agente] ||= []).push(c));
  const agents = Object.entries(allByAgent).filter(([, cs]) => cs.length >= 3).map(([a]) => a);
  const PIL = ['#4A7B9D','#738D84','#C0574A','#6B9E6E','#8FA89F','#20281B','#C9A24B','#9333ea'];

  // datasets per agent
  const datasets = agents.map((agent, i) => {
    const color = PIL[i % PIL.length];
    return {
      label: agent,
      data: weekKeys.map(wk => pxiOf(dated.filter(c => c.weekStart === wk && c.agente === agent))),
      borderColor: color, backgroundColor: color + '22',
      tension: .3, spanGaps: true, pointRadius: 4, pointHoverRadius: 6, borderWidth: 2,
    };
  });
  // clinic average line
  datasets.push({
    label: 'Promedio clínica',
    data: weekKeys.map(wk => pxiOf(dated.filter(c => c.weekStart === wk))),
    borderColor: '#20281B', borderDash: [6, 4], borderWidth: 2.5,
    pointRadius: 3, tension: .3, spanGaps: true,
  });

  destroyChart('chartEvolution');
  charts.chartEvolution = new Chart(document.getElementById('chartEvolution'), {
    type: 'line',
    data: { labels: weekLabels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, boxWidth: 12, color: '#3D4E36' } } },
      scales: {
        y: { min: 0, max: 100, ticks: { callback: v => v + '%', font: { size: 11 }, color: '#3D4E36' }, grid: { color: '#E8E7D8' } },
        x: { ticks: { font: { size: 11 }, color: '#3D4E36' }, grid: { display: false } },
      },
    },
  });

  // per-agent trend cards
  document.getElementById('trendGrid').innerHTML = agents.map(agent => {
    const series = weekKeys.map(wk => ({
      label: weekLabelYr(wk),
      pxi: pxiOf(dated.filter(c => c.weekStart === wk && c.agente === agent)),
    })).filter(p => p.pxi !== null);

    if (!series.length) return '';
    const lastTwo = series.slice(-2);
    const current = series[series.length - 1].pxi;
    const delta = lastTwo.length === 2 ? current - lastTwo[0].pxi : null;
    const st = trendStatus(delta);
    const sparks = series.slice(-8).map(p => {
      const h = Math.max(8, Math.round(p.pxi * 0.5));
      const c = p.pxi >= 80 ? 'var(--ok)' : p.pxi >= 60 ? 'var(--warn)' : 'var(--bad)';
      return `<div class="spark" style="height:${h}px;background:${c}" title="Sem ${p.label}: ${p.pxi}"></div>`;
    }).join('');

    return `<div class="fi-card trend-card">
      <div class="trend-head">
        <div class="trend-agent">${agent}</div>
        <div class="trend-badge ${st.cls}">${st.arrow} ${st.label}</div>
      </div>
      <div class="trend-now">
        <span class="trend-pxi ${pxiClass(current)}">${current}</span>
        <span class="trend-meta">PXI esta semana${delta !== null ? ` · ${delta >= 0 ? '+' : ''}${delta} vs. semana previa` : ''}</span>
      </div>
      <div class="spark-row">${sparks}</div>
      <div class="trend-weeks">${series.length} semana${series.length !== 1 ? 's' : ''} con datos</div>
    </div>`;
  }).join('') || '<div class="empty-state">No hay suficientes semanas con datos todavía.</div>';
}

// ── Flags queue ─────────────────────────────────────────────────────────────
const FLAG_META = {
  r1: { label: 'R1 · Frustración no validada', color: '#C0574A' },
  r2: { label: 'R2 · Info clínica sin confirmación', color: '#C0574A' },
  r3: { label: 'R3 · Cotización alto valor sin supervisor', color: '#4A7B9D' },
  r4: { label: 'R4 · Agente >2h en horario laboral', color: '#C9A24B' },
  r5: { label: 'R5 · Pendiente para turno matutino', color: '#738D84' },
};

function fmtWait(min) {
  if (min == null) return '';
  const h = Math.floor(min / 60), m = min % 60;
  return h ? `${h}h ${m}m` : `${m} min`;
}

function renderFlags() {
  // build filter options
  const aSel = document.getElementById('flagAgent');
  if (aSel.options.length <= 1) {
    [...new Set(weekConversations.map(c => c.agente))].sort().forEach(a => {
      const o = document.createElement('option'); o.value = o.textContent = a; aSel.appendChild(o);
    });
  }

  const items = [];
  weekConversations.forEach(c => {
    Object.keys(c.flags).forEach(f => {
      if (c.flags[f]) items.push({ conv: c, flag: f });
    });
  });
  const filtered = items.filter(i =>
    (!flagFilterAgent || i.conv.agente === flagFilterAgent) &&
    (!flagFilterType || i.flag === flagFilterType));

  document.getElementById('flagCount').textContent =
    `${filtered.length} alerta${filtered.length !== 1 ? 's' : ''}`;

  document.getElementById('flagQueue').innerHTML = filtered.length
    ? filtered.map(({ conv, flag }) => `
      <div class="flag-row" style="border-left-color:${FLAG_META[flag].color}">
        <div class="flag-tag" style="background:${FLAG_META[flag].color}">${FLAG_META[flag].label}</div>
        <div class="flag-body">
          <div class="flag-agent">${conv.agente} · conv. ${conv.num}</div>
          <div class="flag-detail">${flagDetail(conv, flag)}</div>
        </div>
        ${conv.url ? `<a class="flag-link" href="${conv.url}" target="_blank" rel="noopener">Abrir en Atom ↗</a>` : ''}
      </div>`).join('')
    : `<div class="empty-state">✓ Sin alertas con estos filtros</div>`;
}

function flagDetail(c, flag) {
  if (flag === 'r1') return 'El paciente expresó frustración y no se detectó una validación emocional.';
  if (flag === 'r2') return 'Se detectó una afirmación clínica fuera de alcance (tasa de éxito, garantía, etc.).';
  if (flag === 'r3') return 'Se compartió un precio y el hilo menciona FIV/in vitro/PGT/DGP/ovodonación — verificar involucramiento de supervisora.';
  if (flag === 'r4') return `El agente tardó más de 2 h en responder a un lead dentro del horario laboral (7am–medianoche)${c.maxWaitWorking ? ` — espera máxima: ${fmtWait(c.maxWaitWorking)}` : ''}.`;
  if (flag === 'r5') return `El lead escribió fuera de horario o cerca del cierre del turno y quedó sin responder — atender en el primer turno matutino.`;
  return '';
}

function setFlagAgent(v) { flagFilterAgent = v; renderFlags(); }
function setFlagType(v) { flagFilterType = v; renderFlags(); }

// ── Human-sampling tab ──────────────────────────────────────────────────────
const SAMPLE_ITEMS = [
  'Uso real del nombre del paciente / personalización genuina',
  'Profundidad de descubrimiento: ≥2 preguntas SPIN abiertas antes de informar',
  'Protocolo de objeciones completo (Reconocer → Validar → Informar → Invitar)',
  'Siguiente paso concreto y fechado al cierre',
  'HubSpot: motivo de cita + estado del deal documentado',
];

function sampleKey(num, idx) { return `pxi_sample_${num}_${idx}`; }

function renderSampling() {
  const byAgent = {};
  weekConversations.forEach(c => (byAgent[c.agente] ||= []).push(c));

  const html = Object.entries(byAgent).map(([agent, convs]) => {
    // deterministic sample of up to 3
    const picks = [...convs].sort((a, b) => a.num.localeCompare(b.num)).filter((_, i) => i % Math.ceil(convs.length / 3) === 0).slice(0, 3);
    const blocks = picks.map(c => `
      <div class="sample-conv">
        <div class="sample-conv-head">
          <span>Conv. ${c.num}</span>
          ${c.url ? `<a href="${c.url}" target="_blank" rel="noopener">Abrir en Atom ↗</a>` : ''}
        </div>
        ${SAMPLE_ITEMS.map((item, i) => {
          const key = sampleKey(c.num, i);
          const checked = localStorage.getItem(key) === '1' ? 'checked' : '';
          return `<label class="sample-item">
            <input type="checkbox" ${checked} onchange="localStorage.setItem('${key}', this.checked ? '1':'0')" />
            <span>${item}</span></label>`;
        }).join('')}
      </div>`).join('');
    return `<div class="fi-card sample-agent">
      <div class="sample-agent-name">${agent}</div>
      ${blocks || '<div class="empty-state">Sin conversaciones para muestrear</div>'}
    </div>`;
  }).join('');

  document.getElementById('samplingGrid').innerHTML = html;
}

// ── Sample / template ────────────────────────────────────────────────────────
function loadSampleData() { autoLoad(true); }

function downloadTemplate() {
  const headers = ['num_conversacion','cliente_csv','contacto','fecha_inicio_gestion','canal','agente',
    'tipificacion','es_venta','tipo','direccion','remitente','contenido','hora','url'];
  const example = [
    ['1001','Laura Méndez','+5215512345678','01/06/2024','WhatsApp','María López','Interesado','No','mensaje','entrante','Laura Méndez','Hola, ¿qué precio tiene la consulta?','01/06/2024 a las 10:00 am','https://atom.fi/conv/1001'],
    ['1001','Laura Méndez','+5215512345678','01/06/2024','WhatsApp','María López','Interesado','No','mensaje','saliente','María López','¡Hola! La consulta de valoración incluye ultrasonido y evaluación, tiene un costo de $1,200.','01/06/2024 a las 10:08 am','https://atom.fi/conv/1001'],
  ];
  const ws = XLSX.utils.aoa_to_sheet([headers, ...example]);
  ws['!cols'] = headers.map(() => ({ wch: 20 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Historial');
  XLSX.writeFile(wb, 'plantilla_export_atom.xlsx');
}

// ── Auto-load embedded data ──────────────────────────────────────────────────
async function autoLoad(force) {
  for (const name of ['datos_historico.csv', 'datos.xlsx', 'datos.csv']) {
    try {
      const res = await fetch(name, { cache: 'no-store' });
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: false });
      ingestWorkbook(wb, name);
      return true;
    } catch (_) {}
  }
  if (force) showToast('No se encontró datos.xlsx ni datos.csv en el repositorio.');
  return false;
}

// ── Toast ────────────────────────────────────────────────────────────────────
function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3500);
}

autoLoad();
