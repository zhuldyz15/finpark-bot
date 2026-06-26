// Finpark — Telegram-бот: оценка диагностик + «Результаты диагностики» (PDF) + договор (DOCX).
// Принимает текст, ссылку (вкл. Google Docs) и файл .docx/.txt.
// Переменные окружения: BOT_TOKEN, ANTHROPIC_API_KEY (обязательны); ACCESS_CODE, MODEL (необязательны).
const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const PDFDocument = require("pdfkit");
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require("docx");

const DATA = JSON.parse(fs.readFileSync(path.join(__dirname, "rubric_data.json"), "utf8"));
const RUBRIC = DATA.RUBRIC, TIPS = DATA.TIPS;
const CONTRACT_TEMPLATE = fs.readFileSync(path.join(__dirname, "contract_template.txt"), "utf8");
const FONT_R = path.join(__dirname, "DejaVuSans.ttf");
const FONT_B = path.join(__dirname, "DejaVuSans-Bold.ttf");
const CRITERIA = RUBRIC.map(s => "ЭТАП " + s.code + " — " + s.name + ":\n" +
  s.items.map(i => "  " + i.c + " " + i.t + ": " + i.d + " (эталон: " + i.grn + ")").join("\n")).join("\n");

const TOKEN = process.env.BOT_TOKEN;
const KEY = process.env.ANTHROPIC_API_KEY;
const ACCESS_CODE = (process.env.ACCESS_CODE || "").trim();
const MODEL = process.env.MODEL || "claude-haiku-4-5-20251001";
const FRANCHISE_CSV_URL = "https://docs.google.com/spreadsheets/d/1Kt0MZJfzcBxLre4tqR266hGs5yRlLYDDTb0yU28oKEA/export?format=csv";
if (!TOKEN) { console.error("Нет BOT_TOKEN"); process.exit(1); }
if (!KEY) { console.error("Нет ANTHROPIC_API_KEY"); process.exit(1); }
const API = "https://api.telegram.org/bot" + TOKEN;
const allowed = new Set();
const state = {};
const sleep = ms => new Promise(r => setTimeout(r, ms));

const NAVY = "#1F3864", BLUE = "#2E75B6", GREYT = "#595959", LIGHT = "#EAF0FA", GREEN = "#548235";

let franchiseCache = null, franchiseAt = 0;
async function getFranchiseCsv() {
  if (franchiseCache && Date.now() - franchiseAt < 3600000) return franchiseCache;
  try {
    const r = await fetch(FRANCHISE_CSV_URL, { redirect: "follow" });
    const t = await r.text();
    if (t && !/<html/i.test(t)) { franchiseCache = t; franchiseAt = Date.now(); }
  } catch {}
  return franchiseCache || "";
}

async function tg(method, body) {
  const r = await fetch(API + "/" + method, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) });
  return r.json();
}
function esc(s){ return String(s).replace(/[&<>]/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;" }[c])); }
async function send(chatId, text, extra) {
  const parts = []; let cur = "";
  for (const line of text.split("\n")) { if ((cur + line).length > 3800) { parts.push(cur); cur = ""; } cur += line + "\n"; }
  if (cur.trim()) parts.push(cur);
  let last;
  for (let i = 0; i < parts.length; i++) {
    const body = { chat_id: chatId, text: parts[i], parse_mode: "HTML", disable_web_page_preview: true };
    if (i === parts.length - 1 && extra) Object.assign(body, extra);
    last = await tg("sendMessage", body);
  }
  return last;
}
async function sendDocument(chatId, filename, buffer, caption) {
  const fd = new FormData();
  fd.append("chat_id", String(chatId));
  if (caption) fd.append("caption", caption);
  fd.append("document", new Blob([buffer]), filename);
  const r = await fetch(API + "/sendDocument", { method: "POST", body: fd });
  return r.json();
}
function dot(p){ return p>=0.8?"🟢":p>=0.5?"🟡":"🔴"; }
function safeName(s){ return String(s||"Клиент").replace(/[\\/:*?"<>|\n\r]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80) || "Клиент"; }

const WELCOME =
  "<b>Finpark · Бот оценки диагностик</b>\n\n" +
  "Пришлите диагностику любым способом:\n" +
  "• <b>текстом</b> одним сообщением;\n" +
  "• <b>ссылкой</b> (в т.ч. Google Docs — открытый «по ссылке»);\n" +
  "• <b>файлом</b> .docx (Word) или .txt.\n\n" +
  "Я оценю её по 8 этапам продаж (светофор + балл), а затем предложу подготовить " +
  "<b>Результаты диагностики (PDF)</b> и <b>договор</b>.\n\nКоманда: /start — это сообщение.";

// ---------- источники ----------
async function readDocumentText(doc) {
  const f = await tg("getFile", { file_id: doc.file_id });
  if (!f.ok) return null;
  const url = "https://api.telegram.org/file/bot" + TOKEN + "/" + f.result.file_path;
  const name = (doc.file_name || "").toLowerCase();
  const r = await fetch(url);
  if (name.endsWith(".docx")) { const ab = await r.arrayBuffer(); const res = await mammoth.extractRawText({ buffer: Buffer.from(ab) }); return res.value; }
  return await r.text();
}
async function fetchUrlText(rawUrl) {
  let url = rawUrl.trim();
  let m = url.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
  if (m) url = "https://docs.google.com/document/d/" + m[1] + "/export?format=txt";
  else { let s = url.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/); if (s) url = "https://docs.google.com/spreadsheets/d/" + s[1] + "/export?format=csv"; }
  const r = await fetch(url, { redirect: "follow" });
  const ct = (r.headers.get("content-type") || "");
  let t = await r.text();
  if (ct.includes("html")) t = t.replace(/<script[\s\S]*?<\/script>/gi, " ").replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&");
  return t.replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();
}

// ---------- вызов модели ----------
async function callClaude(prompt, maxTokens) {
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "content-type": "application/json", "x-api-key": KEY, "anthropic-version": "2023-06-01" },
    body: JSON.stringify({ model: MODEL, max_tokens: maxTokens || 3000, messages: [{ role: "user", content: prompt }] })
  });
  const d = await r.json();
  if (!r.ok) throw new Error((d.error && d.error.message) || "Ошибка Anthropic API");
  return (d.content && d.content[0] && d.content[0].text) || "";
}
function parseJson(text) { let t = String(text).trim().replace(/^```(?:json)?/i, "").replace(/```$/i, "").trim(); const a = t.indexOf("{"), b = t.lastIndexOf("}"); if (a < 0 || b < 0) throw new Error("нет JSON в ответе"); return JSON.parse(t.slice(a, b + 1)); }
const sval = x => typeof x === "number" ? x : (x && typeof x.s === "number" ? x.s : 0);

const ROP_METHOD =
  "СТИЛЬ РАЗБОРА (как у РОПа Finpark): критикуй прямо, по делу и доказательно — опираясь на конкретные реплики и пропуски в ЭТОМ диалоге, без общих фраз. " +
  "Принципы, которые проверяй жёстко: " +
  "(1) Связь вопрос→презентация: каждый заданный вопрос должен быть отыгран в презентации; лишние вопросы порождают у клиента новые незакрытые вопросы. " +
  "(2) Раскрутка: к базовым вопросам нужны уточняющие; выявить потребности за <10 вопросов невозможно. " +
  "(3) В начале обязателен блок «с кем советоваться/кто принимает решение» — иначе возражение всплывёт в конце. " +
  "(4) Адженда: проговорить цели обеих сторон и вовлечь клиента, держать ценность встречи (не заискивать, не обесценивать). " +
  "(5) Не выдумывать потребность за клиента — фиксировать только реально сказанное; если цель «рост прибыли» — уточнить «от скольки до скольки». " +
  "(6) Презентация адресная под собранные задачи, а не «как вебинар». " +
  "(7) Отработка возражений = возврат к задаче клиента и показ, как продукт её закрывает; на «дорого» спросить «это единственное, что останавливает? на какую сумму ориентировались?». " +
  "(8) Закрытие через фиксацию (предоплата/меньший шаг — финмодель), а не «уходите подумайте». " +
  "(9) Всё взаимосвязано: слабый старт бьёт в финале. Мышление — чистая прибыль, не оборот. ";

const REC_NOTE = "ВАЖНО: пункт 1.1 (встреча на записи) ВСЕГДА считай выполненным (балл 2) — раз есть текст диагностики, запись велась; не считай это ошибкой и не упоминай в разборе. ";

async function evaluateText(transcript) {
  const tx = transcript.slice(0, 55000);
  const rich =
    "Ты — РОП (руководитель отдела продаж) Finpark. Разбери ТЕКСТ ДИАГНОСТИКИ как наставник партнёра-франчайзи, строго по чек-листу. " + ROP_METHOD + REC_NOTE +
    "\nЗадачи: 1) Для КАЖДОГО пункта чек-листа поставь балл 0/1/2 (2 — полно, 1 — частично, 0 — нет/отсутствует), ТОЛЬКО по фактам из текста. " +
    "2) По КАЖДОМУ из 8 этапов: до 2 КОНКРЕТНЫХ ошибок (по реальным репликам/пропускам диалога, коротко) и до 2 «как надо было» (с примером формулировки/вопроса, коротко). " +
    "3) Дай 3 сильные стороны и короткий вердикт.\n\nКРИТЕРИИ:\n" + CRITERIA +
    "\n\nТЕКСТ ДИАГНОСТИКИ:\n" + tx +
    '\n\nВерни СТРОГО валидный JSON без markdown и без лишнего текста. Баллы — числами. Формат: {"scores":{"1.1":2,"1.2":1, ...все коды...},"stages":{"1":{"mistakes":["..."],"howto":["..."]}, ..."8":{...}},"strengths":["..","..",".."],"summary":"вердикт"}';
  let res;
  try { res = parseJson(await callClaude(rich, 8000)); }
  catch (e) {
    const lite =
      "Оцени ТЕКСТ ДИАГНОСТИКИ по чек-листу Finpark. Для каждого пункта балл 0/1/2 по фактам из текста. " + REC_NOTE + "Дай 3 сильные стороны и короткий вердикт.\n\nКРИТЕРИИ:\n" + CRITERIA +
      "\n\nТЕКСТ:\n" + tx +
      '\n\nВерни СТРОГО JSON без markdown: {"scores":{"1.1":2, ...все коды...},"strengths":["..","..",".."],"summary":".."}';
    res = parseJson(await callClaude(lite, 3000));
  }
  if (res && res.scores) res.scores["1.1"] = 2; // запись велась — раз есть текст
  return res;
}

async function extractPrice(transcript) {
  try {
    const t = await callClaude("Найди в тексте диагностики ежемесячную цену (тенге), которую партнёр озвучил клиенту. Верни ТОЛЬКО число без пробелов и символов (например 970000). Если цены нет — верни 0.\n\nТЕКСТ:\n" + transcript.slice(0, 55000), 30);
    const n = (t.match(/\d[\d\s ]{3,}/) || [""])[0].replace(/\D/g, "");
    return n && n.length >= 4 ? n : null;
  } catch { return null; }
}

function buildReport(p) {
  const sc = p.scores || {}; let itog = 0; const stages = [];
  for (const s of RUBRIC) {
    let sum = 0;
    for (const it of s.items) { sum += Math.max(0, Math.min(2, Math.round(sval(sc[it.c])))); }
    const pct = sum / (s.items.length * 2); itog += pct * (s.weight / 100);
    stages.push({ code: s.code, name: s.name, weight: s.weight, pct });
  }
  const growth = [...stages].sort((a, b) => a.pct - b.pct).slice(0, 3).filter(x => x.pct < 0.9);
  const stat = itog>=0.9?"ЭТАЛОН — образец для обучения":itog>=0.75?"ХОРОШО — отшлифовать детали":itog>=0.5?"РАБОЧИЙ УРОВЕНЬ — усилить западающие этапы":"КРИТИЧНО — нужен разбор с тренером";
  let t = "<b>📊 Оценка диагностики</b>\n\n<b>Итог: " + dot(itog) + " " + Math.round(itog*100) + "%</b> — " + stat + "\n";
  if (p.summary) t += "<i>" + esc(p.summary) + "</i>\n";
  t += "\n<b>Светофор по этапам:</b>\n";
  for (const st of stages) t += dot(st.pct) + " " + st.code + ". " + esc(st.name) + " — <b>" + Math.round(st.pct*100) + "%</b> <i>(вес " + st.weight + "%)</i>\n";
  t += "\n<b>💪 Сильные стороны:</b>\n" + ((p.strengths||[]).map(s => "• " + esc(s)).join("\n") || "—") + "\n";
  t += "\n<b>📈 Приоритетные точки роста:</b>\n" + (growth.map(g => dot(g.pct) + " Этап " + g.code + ". " + esc(g.name) + " (" + Math.round(g.pct*100) + "%)").join("\n") || "Все этапы в норме 🟢") + "\n";
  if (growth.length) {
    t += "\n<b>🛠 Что усилить (конкретно):</b>\n";
    for (const g of growth) {
      const sd = (p.stages && p.stages[g.code]) || {};
      const ms = (sd.mistakes || []).slice(0, 2), hw = (sd.howto || []).slice(0, 1);
      t += "<b>Этап " + g.code + ". " + esc(g.name) + "</b>\n";
      ms.forEach(m => t += "✗ " + esc(m) + "\n");
      hw.forEach(h => t += "→ " + esc(h) + "\n");
      if (!ms.length && !hw.length) (TIPS[g.code] || []).slice(0, 1).forEach(x => t += "• " + esc(x) + "\n");
    }
    t += "\n<i>Полный разбор с примерами «как надо» — кнопка «Детальный разбор (PDF)».</i>";
  }
  return t;
}

// ---------- извлечение данных для документа ----------
async function extractKp(transcript) {
  const prompt =
    "Из ТЕКСТА ДИАГНОСТИКИ извлеки данные для документа «Результаты диагностики» Finpark. " +
    "Боли (pains) и решения (solution) формулируй индивидуально и конкретно под этого клиента (5-6 пунктов, по фактам из текста). " +
    "Цену бери ту, что партнёр озвучил клиенту в тексте (ежемесячный платёж в тенге, только цифры, например 970000); если цена не названа — null.\n\n" +
    "ТЕКСТ:\n" + transcript.slice(0, 60000) +
    '\n\nВерни СТРОГО JSON без markdown: {"company":"название или Клиент","niche":"ниша","pains":["..6 болей.."],"goals":["цель1","цель2"],"solution":["как Finpark закроет боль1","..6.."],"price":"970000" или null}';
  return parseJson(await callClaude(prompt, 2000));
}

function fmtPrice(p) { const n = String(p).replace(/\D/g, ""); return n ? n.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + " ₸" : ""; }

// ---------- «Результаты диагностики» (PDF) ----------
function buildKpPdf(kp) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margins: { top: 55, bottom: 60, left: 55, right: 55 }, bufferPages: true });
    doc.registerFont("R", FONT_R); doc.registerFont("B", FONT_B);
    const chunks = []; doc.on("data", c => chunks.push(c)); doc.on("end", () => resolve(Buffer.concat(chunks))); doc.on("error", reject);
    const M = 55, PW = doc.page.width, CW = PW - 2 * M, BOT = () => doc.page.height - 60;
    const company = kp.company && kp.company !== "Клиент" ? kp.company : "вашей компании";
    const ensure = h => { if (doc.y + h > BOT()) doc.addPage(); };
    const H = t => { ensure(50); const y = doc.y; doc.roundedRect(M, y, CW, 28, 4).fill(NAVY); doc.fillColor("#FFFFFF").font("B").fontSize(13).text(t, M + 12, y + 7, { width: CW - 24 }); doc.y = y + 28; doc.moveDown(0.5); doc.fillColor("#000000"); };
    const P = (t, o = {}) => { doc.font(o.b ? "B" : "R").fontSize(o.fs || 11).fillColor(o.color || "#222222"); ensure(doc.heightOfString(t, { width: CW }) + 4); doc.text(t, M, doc.y, { width: CW, align: o.align || "left", paragraphGap: o.gap == null ? 4 : o.gap }); };
    const BL = (arr, check) => { (arr || []).forEach(t => { doc.font("R").fontSize(11); const hh = doc.heightOfString(t, { width: CW - 18 }); ensure(hh + 4); const y = doc.y; doc.font("B").fillColor(check ? GREEN : BLUE).fontSize(11).text(check ? "✓" : "•", M, y, { width: 14 }); doc.font("R").fillColor("#222222").text(t, M + 16, y, { width: CW - 16, paragraphGap: 5 }); }); };
    const statGrid = items => {
      const cols = 3, gap = 10, cardW = (CW - gap * (cols - 1)) / cols, cardH = 58;
      let i = 0;
      while (i < items.length) {
        ensure(cardH + gap);
        const rowY = doc.y;
        for (let col = 0; col < cols && i < items.length; col++, i++) {
          const x = M + col * (cardW + gap);
          doc.roundedRect(x, rowY, cardW, cardH, 4).fillAndStroke(LIGHT, "#D5E0F0");
          doc.fillColor(BLUE).font("B").fontSize(15).text(items[i][0], x + 4, rowY + 9, { width: cardW - 8, align: "center" });
          doc.fillColor(GREYT).font("R").fontSize(8.5).text(items[i][1], x + 6, rowY + 33, { width: cardW - 12, align: "center" });
        }
        doc.y = rowY + cardH + gap;
      }
      doc.x = M; doc.fillColor("#000000");
    };
    const table = (rows, widths, headFill) => {
      const pad = 5;
      rows.forEach((row, ri) => {
        doc.font(ri === 0 ? "B" : "R").fontSize(ri === 0 ? 10 : 9.5);
        let rowH = 0;
        row.forEach((c, ci) => { const h = doc.heightOfString(String(c), { width: widths[ci] - 2 * pad }); if (h > rowH) rowH = h; });
        rowH += 2 * pad;
        ensure(rowH);
        let x = M; const y = doc.y;
        row.forEach((c, ci) => {
          if (ri === 0) doc.rect(x, y, widths[ci], rowH).fill(headFill || NAVY);
          else doc.rect(x, y, widths[ci], rowH).stroke("#DDDDDD");
          doc.fillColor(ri === 0 ? "#FFFFFF" : "#222222").font(ri === 0 ? "B" : "R").fontSize(ri === 0 ? 10 : 9.5)
            .text(String(c), x + pad, y + pad, { width: widths[ci] - 2 * pad });
          x += widths[ci];
        });
        doc.y = y + rowH;
      });
      doc.moveDown(0.5); doc.fillColor("#000000");
    };

    // Обложка
    doc.rect(0, 0, PW, doc.page.height).fill(NAVY);
    doc.fillColor("#FFFFFF").font("B").fontSize(20).text("FINPARK", M, 90, { characterSpacing: 3 });
    doc.font("B").fontSize(34).text("Результаты диагностики", M, 250, { width: CW });
    doc.font("R").fontSize(16).fillColor("#BCD3EE").text("Финансовая система для бизнеса под ключ", M, 320, { width: CW });
    doc.font("R").fontSize(16).fillColor("#FFFFFF").text("для " + company, M, 360, { width: CW });
    doc.font("R").fontSize(11).fillColor("#9FB6D6").text("Онлайн-сопровождение для предпринимателей · " + new Date().toLocaleDateString("ru-RU"), M, doc.page.height - 90, { width: CW });

    doc.addPage(); doc.fillColor("#000000");

    H("О Finpark");
    statGrid([["7+ лет", "на рынке консалтинга"], ["300+", "клиентов по миру"], ["93%", "второй договор"], ["1+ млрд ₸", "прибыли клиентам"], ["923+ млн ₸", "сберегли клиентам"], ["215+ млн ₸", "высвободили в бизнесе"]]);
    P("Работаем с клиентами из 8+ стран: Казахстан, Россия, СНГ, США, ОАЭ, Германия.", { color: GREYT, fs: 10 });
    doc.moveDown(0.4);

    H("Об основательнице компании");
    P("Шанайбаева Жулдыз Абзаловна", { b: true, fs: 12, color: NAVY });
    BL(["Предприниматель и эксперт с опытом в финансах 10+ лет", "4 года в Big4 — руководство аудитом национальных и иностранных компаний", "2+ млрд тенге привлечённых инвестиций", "Спикер международных площадок (BigMoney и др.)", "Квалификация ACCA: 24 место в мире и 1 в Казахстане по экзамену ACCA AAA"]);
    doc.moveDown(0.4);

    H("Проблема: бизнес без управленческого учёта теряет деньги");
    BL(["Открывают новые направления, думая, что увеличат доход", "Сталкиваются с воровством денег и продукции", "Кассовые разрывы — не хватает на зарплаты, дивиденды, поставщиков", "Управляют вслепую, опираясь на ощущения, а не на данные", "Не понимают, почему нет прибыли и как расти дальше"]);
    doc.moveDown(0.4);

    H("Ваша ситуация сейчас");
    BL(kp.pains || []);
    doc.moveDown(0.4);

    H("Решение: финансовый директор на аутсорсе Finpark");
    BL((kp.solution && kp.solution.length ? kp.solution : ["Строит управленческую отчётность и контролирует показатели", "Управляет денежными потоками и финансовой безопасностью", "Выявляет причины потерь и точки роста", "Приводит компанию к конкретным финансовым результатам"]), true);
    P("Мы не просто считаем цифры — мы делаем собственника сильнее.", { color: GREYT, fs: 10 });
    doc.moveDown(0.4);

    H("Чем Finpark отличается от других");
    table([["Другие", "Finpark"], ["Показывают прошлое, считают постфактум", "Показывает точки роста: где теряется прибыль и как исправить"], ["Холодные таблицы и термины, тревога", "Язык собственника, даёт контроль и спокойствие"], ["Дают таблицы, непонятно что делать", "Находит потерянные деньги и превращает в план действий"], ["Не отвечают за результат", "Отвечает за внедрение и результат — ваши дивиденды"]], [CW / 2, CW / 2]);

    H("Наши обязательства перед компанией");
    BL(["Рассчитываем прибыльность: прибыль в ОПиУ без НДС, маржинальность по направлениям, прибыль 3 месяца подряд", "Способствуем росту капитала: управленческий баланс без дыр, ежемесячный рост собственного капитала", "Обеспечиваем финансовую безопасность: баланс капитала 50/50–70/30, контроль дебиторки и кредиторки", "Планируем чистую прибыль с точностью до 80%: отклонение факта от плана не более 20%"], true);
    doc.moveDown(0.4);

    H("Наши специалисты");
    BL(["Финдиректора с опытом в международных компаниях «Большой четвёрки» (Big-4)", "Высшее образование в финансах и бухучёте + повышение квалификации", "Регулярный ассесмент компетенций и подбор сильного состава", "Контроль со стороны руководителей отделов консалтинга и контроля качества", "Аудит каждого звонка финдиректора с клиентом; командное сопровождение"]);
    doc.moveDown(0.4);

    H("Работаем с разными нишами");
    BL(["Общепит: оборачиваемость столов, food/labor cost, калькуляция и анализ меню", "Онлайн-бизнес: CAC, LTV, ROMI, конверсия в оплату, маржинальность продукта", "Производство: себестоимость, эффективность, складские остатки", "Товарный бизнес: оборачиваемость, ABC/XYZ анализ, закуп vs цена, товарные потери"]);
    doc.moveDown(0.4);

    H("Финдиректор Finpark vs штатный");
    table([["Критерий", "Finpark", "Штатный"], ["Оплата", "За фактически оказанные услуги", "Оклад + налоги + льготы"], ["Квалификация", "Команда экспертов из отраслей", "Один специалист"], ["Гибкость", "По мере необходимости", "Постоянная занятость"], ["Ответственность", "За результат", "На собственнике"]], [CW * 0.22, CW * 0.42, CW * 0.36]);

    H("Реальная экономика выбора");
    table([["", "Штатный финдиректор", "Финдиректор Finpark"], ["Затраты в месяц", "1 705 074 ₸ (оклад + налоги)", "от 400 000 – 600 000 ₸"], ["Подтверждение расходов", "Сложно / не всегда", "Прозрачный учёт"], ["Экономия", "—", "до 735 074 ₸/мес · 13,2 млн ₸/год"]], [CW * 0.24, CW * 0.4, CW * 0.36], GREYT);

    H("Что делает финдиректор Finpark");
    P("1. Создаёт архитектуру управленческой отчётности", { b: true, fs: 11, color: NAVY });
    P("Аудит финансовой части → интервью с собственником → финансовая модель и формы отчётов. Итог: ОПиУ, ДДС, баланс и дашборды.");
    P("2. Внедряет финансовый менеджмент", { b: true, fs: 11, color: NAVY });
    P("Стратегическое планирование → тест гипотез → поручения сотрудникам → контроль исполнения. Итог: аналитика под нишу и бюджетирование.");
    doc.moveDown(0.4);

    H("Наше предложение: 6 месяцев сопровождения");
    BL(["Управленческая отчётность: ОПиУ, ДДС, Балансовый отчёт + дашборды", "Финансовая модель на 12 месяцев", "6 стратегических сессий с менеджментом", "До 2 онлайн-встреч с финдиректором в неделю", "24 планёрки по контролю расходов", "Защита цифр перед инвесторами, дорожная карта проекта", "Чат с финдиректором в рабочее время (пн–пт 09:00–18:00)"]);
    doc.moveDown(0.4);

    H("Результат работы с Finpark");
    BL(["Порядок в финансах: внедрены финмодель и отчёты, положительный остаток по операционке", "Контроль расходов: устранены кассовые разрывы, календарь платежей, неликвид обращён в деньги", "Устойчивая прибыль 3 месяца подряд", "Регулярные дивиденды без ущерба для компании", "Данные для управленческих решений на основе твёрдых цифр"], true);
    doc.moveDown(0.4);

    H("Почему нас выбирают — 93% продолжают сотрудничество");
    BL(["Выводим компанию на окупаемость и отвечаем за финансовый результат (методология Finpark, 5+ лет)", "Приучаем мыслить категориями чистой прибыли, а не оборота", "Расходы на финдиректора окупаются: собственник начинает контролировать финансы"]);
    doc.moveDown(0.4);

    H("Кейсы");
    P("Кейс №1 · Услуги — сэкономили 5,2 млн ₸ в год", { b: true, color: NAVY });
    P("Построили ОПиУ по филиалам, выявили убыточный филиал, оптимизировали и закрыли его, перераспределив сотрудников. Снизили убытки на 440 тыс ₸/мес.");
    P("Кейс №2 · Услуги — чистая прибыль выросла на 17 млн ₸", { b: true, color: NAVY });
    P("Построили финмодель, внедрили CRM и пересмотрели KPI менеджеров. Бизнес вышел из убытков, продажи выросли в 2 раза.");
    doc.moveDown(0.4);

    H("Стоимость");
    const m = Number(String(kp.price || "").replace(/\D/g, "")) || 0;
    if (m) {
      const o3 = Math.round(m * 0.959 / 1000) * 1000, o6 = Math.round(m * 0.927 / 1000) * 1000;
      P(fmtPrice(m) + " / месяц", { b: true, fs: 20, color: NAVY });
      doc.moveDown(0.3);
      table([["Вариант оплаты", "Платёж в месяц", "За 6 месяцев"], ["Ежемесячно", fmtPrice(m), fmtPrice(m * 6)], ["При оплате за 3 месяца", fmtPrice(o3), fmtPrice(o3 * 6)], ["При оплате за 6 месяцев", fmtPrice(o6), fmtPrice(o6 * 6)]], [CW * 0.4, CW * 0.3, CW * 0.3]);
      P("Минимальный срок работы по договору — 6 месяцев.", { color: GREYT, fs: 10 });
    } else {
      P("Индивидуальные условия обсуждаются с вашим финансовым директором.", { fs: 13, color: NAVY });
    }
    doc.moveDown(0.6);

    H("Контакты");
    P("Свяжитесь с вашим финансовым директором Finpark, чтобы начать. Не откладывайте внедрение управленческого учёта.", { fs: 12 });

    doc.end();
  });
}

// ---------- Детальный разбор для франчайзи (PDF) ----------
function buildEvalPdf(p) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margins: { top: 55, bottom: 60, left: 55, right: 55 }, bufferPages: true });
    doc.registerFont("R", FONT_R); doc.registerFont("B", FONT_B);
    const chunks = []; doc.on("data", c => chunks.push(c)); doc.on("end", () => resolve(Buffer.concat(chunks))); doc.on("error", reject);
    const M = 55, PW = doc.page.width, CW = PW - 2 * M, BOT = () => doc.page.height - 60;
    const ensure = h => { if (doc.y + h > BOT()) doc.addPage(); };
    const colOf = pct => pct >= 0.8 ? GREEN : pct >= 0.5 ? "#BF8F00" : "#C00000";
    const stWord = pct => pct >= 0.8 ? "норма" : pct >= 0.5 ? "зона роста" : "провал";
    const H = (t, fill) => { ensure(44); const y = doc.y; doc.roundedRect(M, y, CW, 26, 4).fill(fill || NAVY); doc.fillColor("#FFFFFF").font("B").fontSize(12).text(t, M + 12, y + 6, { width: CW - 24 }); doc.y = y + 26; doc.moveDown(0.4); doc.fillColor("#000000"); };
    const P = (t, o = {}) => { doc.font(o.b ? "B" : "R").fontSize(o.fs || 10.5).fillColor(o.color || "#222222"); ensure(doc.heightOfString(t, { width: CW }) + 3); doc.text(t, M, doc.y, { width: CW, paragraphGap: o.gap == null ? 3 : o.gap }); };
    const item = (mark, color, t) => { doc.font("R").fontSize(10.5); const hh = doc.heightOfString(t, { width: CW - 18 }); ensure(hh + 3); const y = doc.y; doc.font("B").fillColor(color).text(mark, M, y, { width: 14 }); doc.font("R").fillColor("#222222").text(t, M + 16, y, { width: CW - 16, paragraphGap: 4 }); };
    const stLine = s => { ensure(16); const y = doc.y; doc.font("B").fillColor(colOf(s.pct)).fontSize(11).text("●", M, y, { width: 12 }); doc.font("R").fillColor("#222222").fontSize(10.5).text("Этап " + s.code + ". " + s.name + " — " + Math.round(s.pct * 100) + "% (вес " + s.weight + "%)", M + 16, y, { width: CW - 16 }); };

    const sc = p.scores || {}; let itog = 0; const stages = [];
    for (const s of RUBRIC) { let sum = 0; for (const it of s.items) { sum += Math.max(0, Math.min(2, Math.round(sval(sc[it.c])))); } const pct = sum / (s.items.length * 2); itog += pct * (s.weight / 100); stages.push({ code: s.code, name: s.name, weight: s.weight, pct }); }
    const stat = itog >= 0.9 ? "Эталон" : itog >= 0.75 ? "Хорошо" : itog >= 0.5 ? "Рабочий уровень" : "Критично";

    doc.rect(0, 0, PW, doc.page.height).fill(NAVY);
    doc.fillColor("#FFFFFF").font("B").fontSize(20).text("FINPARK", M, 90, { characterSpacing: 3 });
    doc.font("B").fontSize(32).text("Разбор диагностики", M, 250, { width: CW });
    doc.font("R").fontSize(15).fillColor("#BCD3EE").text("Внутренний разбор для партнёра-франчайзи", M, 315, { width: CW });
    doc.font("B").fontSize(40).fillColor("#FFFFFF").text(Math.round(itog * 100) + "% · " + stat, M, 385);
    doc.font("R").fontSize(11).fillColor("#9FB6D6").text(new Date().toLocaleDateString("ru-RU"), M, doc.page.height - 90);
    doc.addPage(); doc.fillColor("#000000");

    H("Итог по этапам");
    stages.forEach(stLine);
    doc.moveDown(0.4);
    if (p.summary) { H("Вердикт"); P(p.summary); doc.moveDown(0.2); }
    if (p.strengths && p.strengths.length) { H("Сильные стороны"); p.strengths.forEach(x => item("✓", GREEN, x)); doc.moveDown(0.2); }

    stages.forEach(s => {
      const sd = (p.stages && p.stages[s.code]) || {};
      H("Этап " + s.code + ". " + s.name + "  —  " + Math.round(s.pct * 100) + "% · " + stWord(s.pct), colOf(s.pct) === GREEN ? "#3F6B27" : colOf(s.pct) === "#C00000" ? "#8A1B1B" : "#8A6A00");
      if ((sd.mistakes || []).length) { P("Ошибки:", { b: true, color: "#A01B1B" }); (sd.mistakes || []).forEach(m => item("✗", "#C00000", m)); }
      else P("Серьёзных ошибок не выявлено.", { color: GREYT });
      if ((sd.howto || []).length) { doc.moveDown(0.15); P("Как надо было:", { b: true, color: NAVY }); (sd.howto || []).forEach(h => item("→", BLUE, h)); }
      doc.moveDown(0.3);
    });

    doc.end();
  });
}

// ---------- Договор ----------
async function contractMissing(requisites, price) {
  const franchise = await getFranchiseCsv();
  const prompt =
    "Проверь, достаточно ли данных, чтобы заполнить договор оказания услуг БЕЗ ПРОЧЕРКОВ. " +
    "Обязательные поля: номер договора; по ЗАКАЗЧИКУ и по ИСПОЛНИТЕЛЮ — наименование/ИП, БИН/ИИН, юр.адрес, в лице (ФИО+должность), основание (устав/свидетельство/доверенность), банковские реквизиты; номер и дата договора франшизы исполнителя (его нужно взять из таблицы франшиз по наименованию/БИН/ФИО исполнителя). " +
    "ЦЕНУ НЕ спрашивай — она берётся из диагностики (сейчас: " + (price || "не указана") + "). " +
    "Сопоставь РЕКВИЗИТЫ с обязательными полями и верни список того, чего НЕ ХВАТАЕТ, короткими вопросами на русском. Если франшиза исполнителя не найдена в таблице — добавь вопрос про номер и дату договора франшизы.\n\n" +
    "ТАБЛИЦА ФРАНШИЗ (CSV):\n" + (franchise ? franchise.slice(0, 12000) : "(недоступна)") +
    "\n\nРЕКВИЗИТЫ:\n" + requisites +
    '\n\nВерни СТРОГО JSON без markdown: {"questions":["вопрос1","вопрос2", ...]} (пустой массив, если всё есть).';
  try { const j = parseJson(await callClaude(prompt, 1200)); return Array.isArray(j.questions) ? j.questions : []; }
  catch { return []; }
}

async function buildContractDocx(requisites, price) {
  const franchise = await getFranchiseCsv();
  const prompt =
    "Заполни ПРОПУСКИ в договоре (отмечены подчёркиваниями) данными из РЕКВИЗИТОВ. " +
    "Стоимость услуг за 6 месяцев и ежемесячный платёж рассчитай из ежемесячной цены: ежемесячно = " + (price || "указанной в реквизитах") + " тенге, за 6 месяцев = ежемесячно × 6. " +
    "Везде, где требуется «(сумма прописью)», впиши сумму словами на русском. Дату договора поставь сегодняшней, если номер договора не задан — оставь «№ ___». " +
    "В пункте про «Договор Комплексной Предпринимательской Лицензии (Франчайзинг) №__ от __» подставь НОМЕР И ДАТУ договора франшизы ИСПОЛНИТЕЛЯ из ТАБЛИЦЫ ФРАНШИЗ ниже — найди строку по наименованию/БИН/ФИО исполнителя и возьми значение из колонки «Договор франшизы Номер и Дата». " +
    "Сохрани ВЕСЬ текст и структуру договора без изменений, только заполни пропуски.\n\n" +
    "ТАБЛИЦА ФРАНШИЗ (CSV):\n" + (franchise ? franchise.slice(0, 12000) : "(недоступна)") +
    "\n\nРЕКВИЗИТЫ:\n" + requisites + "\n\nШАБЛОН ДОГОВОРА:\n" + CONTRACT_TEMPLATE +
    "\n\nВерни ТОЛЬКО полный заполненный текст договора, без пояснений и markdown.";
  const filled = await callClaude(prompt, 8000);
  const paras = filled.split("\n").map(line => {
    const t = line.replace(/\s+$/, "");
    if (!t.trim()) return new Paragraph({ spacing: { after: 60 } });
    const isHead = /^\d+\.\s/.test(t.trim()) || /^[А-ЯЁ\s№".]{8,}$/.test(t.trim()) || /^(ЗАКАЗЧИК|ИСПОЛНИТЕЛЬ)/.test(t.trim());
    const center = /ДОГОВОР ОКАЗАНИЯ УСЛУГ/.test(t) || /^г\.\s?Астана/.test(t.trim());
    return new Paragraph({ alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { after: 100 },
      children: [new TextRun({ text: t, bold: !!isHead || center, size: 22, font: "Times New Roman" })] });
  });
  const doc = new Document({
    styles: { default: { document: { run: { font: "Times New Roman", size: 22 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1418 } } }, children: paras }]
  });
  return await Packer.toBuffer(doc);
}

const OFFER_KB = { reply_markup: { inline_keyboard: [
  [{ text: "🧑‍🏫 Детальный разбор для партнёра (PDF)", callback_data: "eval" }],
  [{ text: "📑 Результаты диагностики для клиента (PDF)", callback_data: "kp" }],
  [{ text: "📄 Подготовить договор", callback_data: "contract" }]
] } };

const GUIDE_ZAKAZCHIK =
  "📄 Шаг 1 из 2 — реквизиты <b>ЗАКАЗЧИКА (клиента)</b>. Пришлите одним сообщением:\n\n" +
  "• Наименование / ИП:\n• БИН/ИИН:\n• Юр. адрес:\n• В лице (ФИО, должность):\n• Действует на основании (устав/свид-во):\n• Банковские реквизиты (банк, БИК, IBAN):";
const GUIDE_ISPOLNITEL =
  "📄 Шаг 2 из 2 — реквизиты <b>ИСПОЛНИТЕЛЯ (вас)</b>. Пришлите одним сообщением:\n\n" +
  "• Наименование / ИП:\n• БИН/ИИН:\n• Юр. адрес:\n• В лице (ФИО, должность):\n• Действует на основании:\n• Банковские реквизиты:\n\n" +
  "(Сумму договора возьму из диагностики, номер договора франшизы — из таблицы по вам. Чего не хватит — спрошу отдельными сообщениями.)";

async function generateKp(chatId, transcript, priceOverride) {
  const wait = await tg("sendMessage", { chat_id: chatId, text: "📑 Готовлю «Результаты диагностики»…" });
  try {
    const kp = await extractKp(transcript);
    if (priceOverride) kp.price = priceOverride.replace(/\D/g, "");
    if (!kp.price) {
      state[chatId] = state[chatId] || {}; state[chatId].stage = "await_kp_price"; state[chatId].diagnostic = transcript;
      await send(chatId, "В диагностике не нашёл озвученную цену. Напишите ежемесячную цену, которую вы озвучили клиенту (например: 970000).");
      return;
    }
    const buf = await buildKpPdf(kp);
    const fname = "Результаты диагностики Finpark - " + safeName(kp.company) + ".pdf";
    await sendDocument(chatId, fname, buf, "Результаты диагностики для: " + (kp.company || "клиента"));
    if (state[chatId]) state[chatId].stage = "idle";
  } catch (e) {
    await send(chatId, "⚠️ Не удалось собрать документ: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

async function generateEvalPdf(chatId) {
  const st = state[chatId];
  if (!st || !st.analysis) { await send(chatId, "Сначала пришлите диагностику для оценки."); return; }
  const wait = await tg("sendMessage", { chat_id: chatId, text: "🧑‍🏫 Готовлю детальный разбор…" });
  try {
    const buf = await buildEvalPdf(st.analysis);
    await sendDocument(chatId, "Разбор диагностики Finpark.pdf", buf, "Детальный разбор для партнёра: ошибки по этапам и как надо было.");
  } catch (e) {
    await send(chatId, "⚠️ Не удалось собрать разбор: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

async function generateContract(chatId) {
  const st = state[chatId] || {};
  const wait = await tg("sendMessage", { chat_id: chatId, text: "📄 Готовлю договор…" });
  try {
    const buf = await buildContractDocx(st.creq || "", st.price || null);
    await sendDocument(chatId, "Finpark_Dogovor.docx", buf, "Договор готов. Проверьте реквизиты, номер франшизы и суммы перед подписанием.");
    state[chatId].stage = "idle";
  } catch (e) {
    await send(chatId, "⚠️ Не удалось собрать договор: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

function buildCreq(st) {
  return "ЗАКАЗЧИК (клиент):\n" + (st.zakazchik || "") +
    "\n\nИСПОЛНИТЕЛЬ (партнёр):\n" + (st.ispolnitel || "") +
    "\n\nЕжемесячная цена (из диагностики): " + (st.price ? st.price + " тенге" : "не определена") +
    (st.followups ? "\n\nДополнительно от партнёра:\n" + st.followups : "");
}

// после сбора реквизитов — проверяем, чего не хватает, и докручиваем отдельными сообщениями
async function processContractInfo(chatId) {
  const st = state[chatId];
  st.creq = buildCreq(st);
  const wait = await tg("sendMessage", { chat_id: chatId, text: "🔎 Проверяю данные договора…" });
  let qs = [];
  try { qs = await contractMissing(st.creq, st.price); } catch {}
  if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{});
  st.rounds = (st.rounds || 0) + 1;
  if (qs.length && st.rounds < 4) {
    st.stage = "await_contract_followup";
    await send(chatId, "Ещё нужно уточнить — ответьте следующими сообщениями:\n\n" + qs.map(q => "• " + esc(q)).join("\n") + "\n\nИли напишите <b>генерируй</b>, чтобы выпустить как есть.");
  } else {
    await generateContract(chatId);
  }
}

async function handle(upd) {
  if (upd.callback_query) {
    const cq = upd.callback_query; const chatId = cq.message.chat.id;
    tg("answerCallbackQuery", { callback_query_id: cq.id }).catch(()=>{});
    const st = state[chatId];
    if (cq.data === "eval") {
      await generateEvalPdf(chatId);
    } else if (cq.data === "kp") {
      if (!st || !st.diagnostic) { await send(chatId, "Сначала пришлите диагностику для оценки."); return; }
      await generateKp(chatId, st.diagnostic);
    } else if (cq.data === "contract") {
      const s2 = state[chatId] = state[chatId] || {};
      s2.stage = "await_zakazchik"; s2.rounds = 0; s2.zakazchik = null; s2.ispolnitel = null; s2.followups = "";
      if (!s2.price && s2.diagnostic) { try { s2.price = await extractPrice(s2.diagnostic); } catch {} }
      await send(chatId, GUIDE_ZAKAZCHIK);
    }
    return;
  }

  const msg = upd.message; if (!msg) return;
  const chatId = msg.chat.id;
  const text = (msg.text || "").trim();

  if (text === "/start") { state[chatId] = { stage: "idle" }; await send(chatId, WELCOME); if (ACCESS_CODE) await send(chatId, "🔒 Доступ по коду. Пришлите кодовое слово, чтобы начать."); return; }

  if (ACCESS_CODE && !allowed.has(chatId)) {
    if (text === ACCESS_CODE) { allowed.add(chatId); await send(chatId, "✅ Доступ открыт. Пришлите диагностику текстом, ссылкой или файлом."); }
    else await send(chatId, "🔒 Неверный код. Пришлите правильное кодовое слово.");
    return;
  }

  const st = state[chatId];
  if (st && st.stage === "await_kp_price" && text) {
    const price = text.replace(/\D/g, "");
    if (price.length < 4) { await send(chatId, "Похоже, это не сумма. Напишите цену цифрами, например: 970000."); return; }
    st.price = price; await generateKp(chatId, st.diagnostic, price); return;
  }
  if (st && st.stage === "await_zakazchik" && (text || msg.document)) {
    let req = text;
    if (msg.document) { try { req = await readDocumentText(msg.document) || text; } catch {} }
    if (!req || req.length < 8) { await send(chatId, "Пришлите реквизиты заказчика текстом."); return; }
    st.zakazchik = req; st.stage = "await_ispolnitel";
    await send(chatId, GUIDE_ISPOLNITEL); return;
  }
  if (st && st.stage === "await_ispolnitel" && (text || msg.document)) {
    let req = text;
    if (msg.document) { try { req = await readDocumentText(msg.document) || text; } catch {} }
    if (!req || req.length < 8) { await send(chatId, "Пришлите реквизиты исполнителя текстом."); return; }
    st.ispolnitel = req; await processContractInfo(chatId); return;
  }
  if (st && st.stage === "await_contract_followup" && text) {
    if (/^(генерируй|готово|выпускай|давай)/i.test(text)) { await generateContract(chatId); return; }
    st.followups = (st.followups || "") + (st.followups ? "\n" : "") + text;
    await processContractInfo(chatId); return;
  }

  let transcript = "";
  try {
    if (msg.document) {
      const name = (msg.document.file_name || "").toLowerCase();
      if (name.endsWith(".docx") || name.endsWith(".txt")) transcript = await readDocumentText(msg.document) || "";
      else if (name.endsWith(".doc")) { await send(chatId, "Формат .doc не поддерживается. Пересохраните как .docx."); return; }
      else { await send(chatId, "Поддерживаются файлы .docx и .txt, либо текст/ссылка."); return; }
    } else if (/^https?:\/\//i.test(text)) {
      const w0 = await tg("sendMessage", { chat_id: chatId, text: "🔗 Загружаю текст по ссылке…" });
      transcript = await fetchUrlText(text);
      if (w0.ok) tg("deleteMessage", { chat_id: chatId, message_id: w0.result.message_id }).catch(()=>{});
      if (!transcript || transcript.length < 40) { await send(chatId, "По ссылке не удалось получить текст. Проверьте доступ «по ссылке» или пришлите текст/файл."); return; }
    } else { transcript = text; }
  } catch (e) { await send(chatId, "⚠️ Не удалось прочитать материал: " + esc(e.message || String(e))); return; }

  if (!transcript || transcript.trim().length < 40) { await send(chatId, "Пришлите полный текст диагностики (хотя бы несколько реплик) — текстом, ссылкой или файлом .docx/.txt."); return; }

  const wait = await tg("sendMessage", { chat_id: chatId, text: "⏳ Анализирую диагностику по 38 критериям…" });
  try {
    const parsed = await evaluateText(transcript);
    state[chatId] = { stage: "idle", diagnostic: transcript, analysis: parsed };
    await send(chatId, buildReport(parsed));
    await send(chatId, "Нужно подготовить документы по этому клиенту?", OFFER_KB);
  } catch (e) {
    await send(chatId, "⚠️ Ошибка оценки: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

async function main() {
  console.log("Finpark bot v4 запущен. Модель:", MODEL, "| код доступа:", ACCESS_CODE ? "включён" : "выключен");
  await tg("deleteWebhook", { drop_pending_updates: false }).catch(()=>{});
  let offset = 0;
  while (true) {
    try {
      const u = await tg("getUpdates", { offset, timeout: 50, allowed_updates: ["message", "callback_query"] });
      for (const upd of (u.result || [])) { offset = upd.update_id + 1; handle(upd).catch(e => console.error("handle:", e)); }
    } catch (e) { console.error("poll:", e.message); await sleep(2000); }
  }
}
main();
