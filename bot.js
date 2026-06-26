// Finpark — Telegram-бот: оценка диагностик + генерация КП (презентация) и договора.
// Принимает текст, ссылку (вкл. Google Docs) и файл .docx/.txt.
// Переменные окружения: BOT_TOKEN, ANTHROPIC_API_KEY (обязательны); ACCESS_CODE, MODEL (необязательны).
const fs = require("fs");
const path = require("path");
const mammoth = require("mammoth");
const PptxGenJS = require("pptxgenjs");
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require("docx");

const DATA = JSON.parse(fs.readFileSync(path.join(__dirname, "rubric_data.json"), "utf8"));
const RUBRIC = DATA.RUBRIC, TIPS = DATA.TIPS;
const CONTRACT_TEMPLATE = fs.readFileSync(path.join(__dirname, "contract_template.txt"), "utf8");
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

const NAVY = "1F3864", BLUE = "2E75B6", GREYT = "595959", LIGHT = "EAF0FA", GOLD = "BF8F00";

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

const WELCOME =
  "<b>Finpark · Бот оценки диагностик</b>\n\n" +
  "Пришлите диагностику любым способом:\n" +
  "• <b>текстом</b> одним сообщением;\n" +
  "• <b>ссылкой</b> (в т.ч. Google Docs — открытый «по ссылке»);\n" +
  "• <b>файлом</b> .docx (Word) или .txt.\n\n" +
  "Я оценю её по 8 этапам продаж (светофор + балл), а затем предложу подготовить " +
  "<b>КП-презентацию</b> и <b>договор</b>.\n\nКоманда: /start — это сообщение.";

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
function parseJson(text) { const a = text.indexOf("{"), b = text.lastIndexOf("}"); return JSON.parse(text.slice(a, b + 1)); }

async function evaluateText(transcript) {
  const prompt =
    "Ты — методолог отдела продаж Finpark (финдиректор на аутсорсе). Оцени ТЕКСТ ДИАГНОСТИКИ строго по критериям. " +
    "Для КАЖДОГО пункта поставь балл: 2 если выполнен в полном объёме, 1 если частично/поверхностно, 0 если не выполнено или отсутствует. " +
    "Оценивай ТОЛЬКО по тому, что реально есть в тексте. Дай также 3 сильные стороны и 1 короткое резюме.\n\nКРИТЕРИИ:\n" + CRITERIA +
    "\n\nТЕКСТ ДИАГНОСТИКИ:\n" + transcript.slice(0, 60000) +
    '\n\nВерни СТРОГО валидный JSON без markdown: {"scores":{"1.1":{"s":2,"n":"кратко"}, ... все коды ...},"strengths":["..","..",".."],"summary":".."}';
  return parseJson(await callClaude(prompt));
}

function buildReport(p) {
  const sc = p.scores || {}; let itog = 0; const stages = [];
  for (const s of RUBRIC) {
    let sum = 0;
    for (const it of s.items) { let v = sc[it.c] && typeof sc[it.c].s === "number" ? sc[it.c].s : 0; v = Math.max(0, Math.min(2, Math.round(v))); sum += v; }
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
  if (growth.length) { t += "\n<b>🛠 Рекомендации:</b>\n"; for (const g of growth) t += "<b>Этап " + g.code + ". " + esc(g.name) + "</b>\n" + (TIPS[g.code]||[]).map(x => "• " + esc(x)).join("\n") + "\n"; }
  return t;
}

// ---------- КП: извлечение данных ----------
async function extractKp(transcript) {
  const prompt =
    "Из ТЕКСТА ДИАГНОСТИКИ извлеки данные для коммерческого предложения Finpark. " +
    "Боли (pains) и решения (solution) формулируй индивидуально и конкретно под этого клиента (5-6 пунктов, по фактам из текста). " +
    "Цену бери ту, что партнёр озвучил клиенту в тексте (ежемесячный платёж в тенге, только цифры, например 970000); если цена не названа — null.\n\n" +
    "ТЕКСТ:\n" + transcript.slice(0, 60000) +
    '\n\nВерни СТРОГО JSON без markdown: {"company":"название или Клиент","niche":"ниша","pains":["..6 болей.."],"goals":["цель1","цель2"],"solution":["как Finpark закроет боль1","..6.."],"price":"970000" или null}';
  return parseJson(await callClaude(prompt, 2000));
}

function fmtPrice(p) { const n = String(p).replace(/\D/g, ""); return n ? n.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + " ₸" : ""; }

// ---------- КП: сборка презентации (полный набор разделов) ----------
async function buildKpPptx(kp, chatId) {
  const p = new PptxGenJS();
  p.defineLayout({ name: "FP", width: 13.33, height: 7.5 });
  p.layout = "FP";
  const W = 13.33, H = 7.5;
  const company = kp.company && kp.company !== "Клиент" ? kp.company : "вашей компании";

  const header = (s, txt) => {
    s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: W, h: 1.15, fill: { color: NAVY } });
    s.addText(txt, { x: 0.6, y: 0.18, w: W - 1.2, h: 0.8, fontSize: 24, bold: true, color: "FFFFFF", valign: "middle" });
  };
  const bullets = (s, arr, opt = {}) => s.addText((arr || []).map(t => ({ text: t, options: { bullet: { code: opt.check ? "2713" : "2022" }, color: opt.color || "333333", fontSize: opt.fs || 17, paraSpaceAfter: 9 } })), { x: 0.7, y: 1.5, w: W - 1.4, h: H - 2, valign: "top" });
  const note = (s, txt, y) => s.addText(txt, { x: 0.7, y: y || (H - 1), w: W - 1.4, h: 0.7, fontSize: 14, italic: true, color: GREYT });

  // 1. Титул
  let s = p.addSlide(); s.background = { color: NAVY };
  s.addText("FINPARK", { x: 0.7, y: 0.7, w: 6, h: 0.6, fontSize: 22, bold: true, color: "FFFFFF", charSpacing: 4 });
  s.addText("Коммерческое предложение", { x: 0.7, y: 2.4, w: W - 1.4, h: 1, fontSize: 40, bold: true, color: "FFFFFF" });
  s.addText("Финансовая система для бизнеса под ключ", { x: 0.7, y: 3.5, w: W - 1.4, h: 0.7, fontSize: 22, color: "BCD3EE" });
  s.addText("для " + company, { x: 0.7, y: 4.5, w: W - 1.4, h: 0.7, fontSize: 22, italic: true, color: "FFFFFF" });
  s.addText("Онлайн-сопровождение для предпринимателей", { x: 0.7, y: H - 0.9, w: W - 1.4, h: 0.5, fontSize: 14, color: "9FB6D6" });

  // 2. По итогам встречи
  s = p.addSlide(); header(s, "По итогам встречи вы узнаете");
  bullets(s, ["Для чего в компаниях нужен управленческий учёт", "Какие проблемы решает финансовый директор на аутсорсе от Finpark", "В чём преимущества финдиректора на аутсорсе", "Какие у финдиректора обязанности и обязательства перед компанией", "Как изменятся финансовое и стратегическое планирование после внедрения управленческого учёта"]);

  // 3. О Finpark — цифры
  s = p.addSlide(); header(s, "О Finpark");
  const stats = [["7+ лет", "на рынке финансового консалтинга"], ["300+", "клиентов по всему миру"], ["93%", "подписывают второй договор"], ["1+ млрд ₸", "увеличили прибыль клиентов"], ["923+ млн ₸", "сберегли клиентам"], ["215+ млн ₸", "высвободили внутри бизнеса"]];
  stats.forEach((st, i) => {
    const x = 0.7 + (i % 3) * 4.1, y = 1.6 + Math.floor(i / 3) * 2.5;
    s.addShape(p.ShapeType.roundRect, { x, y, w: 3.8, h: 2.2, fill: { color: LIGHT }, line: { color: "D5E0F0" }, rectRadius: 0.08 });
    s.addText(st[0], { x: x + 0.1, y: y + 0.25, w: 3.6, h: 0.9, fontSize: 28, bold: true, color: BLUE, align: "center" });
    s.addText(st[1], { x: x + 0.15, y: y + 1.15, w: 3.5, h: 0.9, fontSize: 14, color: GREYT, align: "center" });
  });
  note(s, "Работаем с клиентами из 8+ стран: Казахстан, Россия, СНГ, США, ОАЭ, Германия.");

  // 4. Об основательнице
  s = p.addSlide(); header(s, "Об основательнице компании");
  s.addText("Шанайбаева Жулдыз Абзаловна", { x: 0.7, y: 1.4, w: W - 1.4, h: 0.6, fontSize: 22, bold: true, color: NAVY });
  bullets(s, ["Предприниматель и эксперт с опытом в финансах 10+ лет", "4 года в Big4 — руководство аудитом национальных и иностранных компаний", "2+ млрд тенге привлечённых инвестиций", "Спикер международных площадок (BigMoney и др.)", "Квалификация ACCA: 24 место в мире и 1 место в Казахстане по экзамену ACCA AAA"], { fs: 16 });
  s.getSlide && 0;
  s.addText("В основе методологии Finpark — международный опыт, адаптированный под местный менталитет.", { x: 0.7, y: H - 1, w: W - 1.4, h: 0.6, fontSize: 14, italic: true, color: GREYT });

  // 5. Проблема (общая)
  s = p.addSlide(); header(s, "Проблема: бизнес без управленческого учёта теряет деньги");
  bullets(s, ["Открывают новые направления, думая, что увеличат доход", "Сталкиваются с воровством денег и продукции", "Имеют кассовые разрывы — не хватает на зарплаты, дивиденды, поставщиков", "Управляют вслепую, опираясь на ощущения, а не на данные", "Не понимают, почему нет прибыли и как расти дальше", "Не зарабатывают, потому что не видят свои дивиденды"], { fs: 16 });

  // 6. Ваша ситуация (ПЕРСОНАЛЬНО)
  s = p.addSlide(); header(s, "Ваша ситуация сейчас");
  bullets(s, kp.pains || [], { fs: 18 });

  // 7. Решение
  s = p.addSlide(); header(s, "Решение: финансовый директор на аутсорсе Finpark");
  bullets(s, (kp.solution && kp.solution.length ? kp.solution : ["Погружается в процессы и строит управленческую отчётность", "Контролирует выполнение финансовых показателей", "Управляет денежными потоками и финансовой безопасностью", "Выявляет причины потерь и точки роста", "Приводит компанию к конкретным финансовым результатам"]), { check: true, fs: 17 });
  note(s, "Мы не просто считаем цифры — мы делаем собственника сильнее.");

  // 8. Чем отличаемся
  s = p.addSlide(); header(s, "Чем Finpark отличается от других");
  s.addTable([
    [{ text: "Другие", options: { bold: true, color: "FFFFFF", fill: GREYT } }, { text: "Finpark", options: { bold: true, color: "FFFFFF", fill: NAVY } }],
    ["Показывают прошлое, считают постфактум", "Показывает точки роста: где теряется прибыль и как это исправить"],
    ["Холодные таблицы и термины, создают тревогу", "Язык собственника, убирает тревогу, даёт контроль"],
    ["Дают таблицы, но непонятно, что делать", "Находит потерянные деньги и превращает в план действий"],
    ["Не отвечают, сработают ли рекомендации", "Отвечает за внедрение системы и результат — ваши дивиденды"]
  ], { x: 0.6, y: 1.5, w: W - 1.2, colW: [6.05, 6.08], fontSize: 14, border: { type: "solid", color: "DDDDDD" }, valign: "middle", rowH: 0.9 });

  // 9. Обязательства
  s = p.addSlide(); header(s, "Наши обязательства перед компанией");
  bullets(s, ["Рассчитываем прибыльность: прибыль в ОПиУ без НДС, маржинальность по направлениям, прибыль 3 месяца подряд", "Способствуем росту капитала: управленческий баланс без дыр, ежемесячный рост собственного капитала", "Обеспечиваем финансовую безопасность: баланс капитала 50/50–70/30, контроль дебиторки и кредиторки", "Планируем чистую прибыль с точностью до 80%: отклонение факта от плана не более 20%"], { check: true, fs: 16 });
  note(s, "«Где больше правды в цифрах, там больше денег» — философия Finpark.");

  // 10. Специалисты
  s = p.addSlide(); header(s, "Наши специалисты");
  bullets(s, ["Финдиректора с опытом в международных компаниях «Большой четвёрки» (Big-4)", "Высшее образование в сфере финансов и бухучёта + системное повышение квалификации", "Регулярный ассесмент компетенций и подбор сильного состава", "Контроль со стороны руководителей отделов консалтинга и контроля качества", "Аудит каждого звонка финдиректора с клиентом", "Командное сопровождение: финдиректор + РОК + менеджер"], { fs: 16 });

  // 11. Ниши
  s = p.addSlide(); header(s, "Работаем с разными бизнес-моделями");
  const niches = [["Общепит", "оборачиваемость столов, food cost / labor cost, калькуляция и анализ меню, загрузка по времени"], ["Онлайн-бизнес", "CAC, LTV, ROMI, конверсия в оплату, маржинальность продукта"], ["Производство", "себестоимость, эффективность производства, складские остатки"], ["Товарный бизнес", "оборачиваемость остатков, ABC/XYZ анализ, закуп vs цена, товарные потери"]];
  niches.forEach((n, i) => {
    const x = 0.7 + (i % 2) * 6.1, y = 1.5 + Math.floor(i / 2) * 2.6;
    s.addShape(p.ShapeType.roundRect, { x, y, w: 5.8, h: 2.3, fill: { color: LIGHT }, line: { color: "D5E0F0" }, rectRadius: 0.06 });
    s.addText(n[0], { x: x + 0.25, y: y + 0.2, w: 5.3, h: 0.5, fontSize: 18, bold: true, color: NAVY });
    s.addText(n[1], { x: x + 0.25, y: y + 0.8, w: 5.3, h: 1.4, fontSize: 13, color: "333333" });
  });

  // 12. Finpark vs штатный
  s = p.addSlide(); header(s, "Финдиректор Finpark vs штатный");
  s.addTable([
    [{ text: "Критерий", options: { bold: true, color: "FFFFFF", fill: NAVY } }, { text: "Finpark", options: { bold: true, color: "FFFFFF", fill: NAVY } }, { text: "Штатный", options: { bold: true, color: "FFFFFF", fill: NAVY } }],
    ["Оплата", "За фактически оказанные услуги", "Полный оклад + налоги + льготы"],
    ["Квалификация", "Команда экспертов из разных отраслей", "Один специалист, ограниченный опыт"],
    ["Гибкость", "Привлечение по мере необходимости", "Постоянное рабочее место и занятость"],
    ["Ответственность", "Берёт ответственность за результат", "Остаётся на собственнике"]
  ], { x: 0.6, y: 1.5, w: W - 1.2, colW: [3, 4.55, 4.58], fontSize: 14, border: { type: "solid", color: "DDDDDD" }, valign: "middle", rowH: 0.85 });

  // 13. Реальная экономика
  s = p.addSlide(); header(s, "Реальная экономика выбора");
  s.addTable([
    [{ text: "", options: { fill: "FFFFFF" } }, { text: "Штатный финдиректор", options: { bold: true, color: "FFFFFF", fill: GREYT } }, { text: "Финдиректор Finpark", options: { bold: true, color: "FFFFFF", fill: NAVY } }],
    ["Затраты в месяц", "1 705 074 ₸ (оклад + налоги + рабочее место)", "от 400 000 – 600 000 ₸"],
    ["Подтверждение расходов", "Сложно / не всегда", "Прозрачный затратный учёт"],
    ["Экономия", "—", "до 735 074 ₸ в месяц · 13,2 млн ₸ в год"]
  ], { x: 0.6, y: 1.5, w: W - 1.2, colW: [3, 4.55, 4.58], fontSize: 14, border: { type: "solid", color: "DDDDDD" }, valign: "middle", rowH: 0.95 });

  // 14. Виды работ
  s = p.addSlide(); header(s, "Что делает финдиректор Finpark");
  s.addText("1. Создаёт архитектуру управленческой отчётности", { x: 0.7, y: 1.5, w: W - 1.4, h: 0.5, fontSize: 17, bold: true, color: NAVY });
  s.addText("Аудит финансовой части → интервью с собственником → финансовая модель и формы отчётов. Итог: ОПиУ, ДДС, баланс и дашборды.", { x: 0.9, y: 2.0, w: W - 1.8, h: 1.1, fontSize: 14, color: "333333" });
  s.addText("2. Внедряет финансовый менеджмент", { x: 0.7, y: 3.4, w: W - 1.4, h: 0.5, fontSize: 17, bold: true, color: NAVY });
  s.addText("Стратегическое планирование → тест гипотез → поручения сотрудникам → контроль исполнения в рамках долгосрочных целей. Итог: аналитика под нишу и бюджетирование.", { x: 0.9, y: 3.9, w: W - 1.8, h: 1.2, fontSize: 14, color: "333333" });

  // 15. Наше предложение
  s = p.addSlide(); header(s, "Наше предложение: 6 месяцев сопровождения");
  bullets(s, ["Управленческая отчётность: ОПиУ, ДДС, Балансовый отчёт + дашборды", "Финансовая модель на 12 месяцев", "6 стратегических сессий с менеджментом", "До 2 онлайн-встреч с финдиректором в неделю", "24 планёрки по контролю расходов", "Защита цифр перед инвесторами, дорожная карта проекта", "Чат с финдиректором в рабочее время (пн–пт 09:00–18:00)"], { fs: 15 });

  // 16. Обучение
  s = p.addSlide(); header(s, "Обучение и развитие");
  s.addText("Для собственника:", { x: 0.7, y: 1.5, w: 5.8, h: 0.5, fontSize: 16, bold: true, color: NAVY });
  s.addText([{ text: "Развитие финансового мышления", options: { bullet: { code: "2022" }, fontSize: 14, paraSpaceAfter: 6 } }, { text: "Принятие решений через цифры", options: { bullet: { code: "2022" }, fontSize: 14, paraSpaceAfter: 6 } }, { text: "Материалы по финансовой грамотности", options: { bullet: { code: "2022" }, fontSize: 14 } }], { x: 0.9, y: 2.0, w: 5.5, h: 2.5, valign: "top", color: "333333" });
  s.addText("Для команды:", { x: 6.9, y: 1.5, w: 5.8, h: 0.5, fontSize: 16, bold: true, color: NAVY });
  s.addText([{ text: "Программа «Сам себе финдир»", options: { bullet: { code: "2022" }, fontSize: 14, paraSpaceAfter: 6 } }, { text: "Базовая финансовая грамотность", options: { bullet: { code: "2022" }, fontSize: 14, paraSpaceAfter: 6 } }, { text: "Работа с управленческой отчётностью", options: { bullet: { code: "2022" }, fontSize: 14 } }], { x: 7.1, y: 2.0, w: 5.5, h: 2.5, valign: "top", color: "333333" });

  // 17. Результат
  s = p.addSlide(); header(s, "Результат работы с Finpark");
  bullets(s, ["Порядок в финансах: внедрены финмодель и управленческие отчёты, положительный остаток по операционке", "Контроль расходов: устранены кассовые разрывы, календарь платежей, неликвид обращён в деньги", "Стратегия развития: устойчивая прибыль 3 месяца подряд", "Регулярные дивиденды в любых обстоятельствах без ущерба для компании", "Данные для управленческих решений на основе твёрдых цифр"], { check: true, fs: 16 });

  // 18. Почему выбирают
  s = p.addSlide(); header(s, "Почему нас выбирают — 93% продолжают сотрудничество");
  bullets(s, ["Выводим компанию на окупаемость и отвечаем за финансовый результат (авторская методология Finpark, 5+ лет)", "Приучаем мыслить категориями чистой прибыли, а не оборота — показываем на цифрах", "Расходы на финдиректора Finpark окупаются: собственник начинает понимать процессы и контролировать финансы"], { fs: 16 });

  // 19. Кейсы
  s = p.addSlide(); header(s, "Кейсы");
  s.addText("Кейс №1 · Услуги — сэкономили 5,2 млн ₸ в год", { x: 0.7, y: 1.45, w: W - 1.4, h: 0.5, fontSize: 16, bold: true, color: NAVY });
  s.addText("Построили ОПиУ с разделением по филиалам, выявили убыточный филиал, оптимизировали и закрыли его, перераспределив сотрудников. Снизили убытки на 440 тыс ₸/мес.", { x: 0.9, y: 1.95, w: W - 1.8, h: 1.2, fontSize: 14, color: "333333" });
  s.addText("Кейс №2 · Услуги — чистая прибыль выросла на 17 млн ₸", { x: 0.7, y: 3.5, w: W - 1.4, h: 0.5, fontSize: 16, bold: true, color: NAVY });
  s.addText("Построили финмодель, внедрили CRM и пересмотрели KPI менеджеров (план продаж, конверсия, средний чек). Бизнес вышел из убытков, продажи выросли в 2 раза.", { x: 0.9, y: 4.0, w: W - 1.8, h: 1.2, fontSize: 14, color: "333333" });

  // 20. Примеры оформления
  s = p.addSlide(); header(s, "Примеры оформления");
  bullets(s, ["Дорожная карта проекта — этапы работы с конкретными сроками", "Финансовая модель — проверка бизнес-идей в цифрах", "Отчёт о движении денежных средств (ДДС)", "Отчёт о прибылях и убытках (ОПиУ)", "Балансовый отчёт — активы и обязательства", "Дашборды и аналитика под вашу нишу"], { fs: 16 });

  // 21. Стоимость (ПЕРСОНАЛЬНО)
  s = p.addSlide(); s.background = { color: LIGHT }; header(s, "Стоимость");
  const m = Number(String(kp.price || "").replace(/\D/g, "")) || 0;
  if (m) {
    const opt3 = Math.round(m * 0.959 / 1000) * 1000, opt6 = Math.round(m * 0.927 / 1000) * 1000;
    s.addText(fmtPrice(m) + " / месяц", { x: 0.7, y: 1.5, w: W - 1.4, h: 1, fontSize: 40, bold: true, color: NAVY });
    s.addTable([
      [{ text: "Вариант оплаты", options: { bold: true, color: "FFFFFF", fill: NAVY } }, { text: "Платёж в месяц", options: { bold: true, color: "FFFFFF", fill: NAVY } }, { text: "За 6 месяцев", options: { bold: true, color: "FFFFFF", fill: NAVY } }],
      ["Ежемесячно", fmtPrice(m), fmtPrice(m * 6)],
      ["При оплате за 3 месяца", fmtPrice(opt3), fmtPrice(opt6 ? opt3 * 6 : m * 6)],
      ["При оплате за 6 месяцев", fmtPrice(opt6), fmtPrice(opt6 * 6)]
    ], { x: 0.7, y: 2.7, w: W - 1.4, colW: [4.6, 3.6, 3.93], fontSize: 15, border: { type: "solid", color: "C9D6EA" }, valign: "middle", rowH: 0.7, fill: { color: "FFFFFF" } });
    s.addText("Минимальный срок работы по договору — 6 месяцев.", { x: 0.7, y: H - 1.1, w: W - 1.4, h: 0.5, fontSize: 14, italic: true, color: GREYT });
  } else {
    s.addText("Индивидуальные условия обсуждаются с вашим финансовым директором.", { x: 0.7, y: 2.5, w: W - 1.4, h: 1, fontSize: 24, color: NAVY });
  }

  // 22. Контакты
  s = p.addSlide(); s.background = { color: NAVY };
  s.addText("Увеличьте финансовые показатели уже сейчас", { x: 0.7, y: 2.4, w: W - 1.4, h: 1.2, fontSize: 30, bold: true, color: "FFFFFF" });
  s.addText("Не откладывайте внедрение управленческого учёта.", { x: 0.7, y: 3.6, w: W - 1.4, h: 0.7, fontSize: 18, color: "BCD3EE" });
  s.addText("Свяжитесь с вашим финансовым директором Finpark для старта.", { x: 0.7, y: 4.3, w: W - 1.4, h: 0.7, fontSize: 16, color: "9FB6D6" });

  const file = path.join("/tmp", "kp_" + chatId + "_" + Date.now() + ".pptx");
  await p.writeFile({ fileName: file });
  const buf = fs.readFileSync(file); fs.unlink(file, () => {});
  return buf;
}

// ---------- Договор ----------
async function buildContractDocx(requisites, price) {
  const franchise = await getFranchiseCsv();
  const prompt =
    "Заполни ПРОПУСКИ в договоре (отмечены подчёркиваниями) данными из РЕКВИЗИТОВ. " +
    "Стоимость услуг за 6 месяцев и ежемесячный платёж рассчитай из ежемесячной цены: ежемесячно = " + (price || "указанной в реквизитах") + " тенге, за 6 месяцев = ежемесячно × 6. " +
    "Везде, где требуется «(сумма прописью)», впиши сумму словами на русском. Дату договора поставь сегодняшней. " +
    "В пункте про «Договор Комплексной Предпринимательской Лицензии (Франчайзинг) №__ от __» подставь НОМЕР И ДАТУ договора франшизы ИСПОЛНИТЕЛЯ из ТАБЛИЦЫ ФРАНШИЗ ниже — найди строку по наименованию/БИН/ФИО исполнителя и возьми значение из колонки «Договор франшизы Номер и Дата» (например «№13 от 02.05.2024»). Если исполнитель в таблице не найден — оставь пропуск. " +
    "Сохрани ВЕСЬ текст и структуру договора без изменений, только заполни пропуски. Если каких-то данных нет — оставь короткий пропуск «____».\n\n" +
    "ТАБЛИЦА ФРАНШИЗ (CSV):\n" + (franchise ? franchise.slice(0, 12000) : "(недоступна)") +
    "\n\nРЕКВИЗИТЫ:\n" + requisites + "\n\nШАБЛОН ДОГОВОРА:\n" + CONTRACT_TEMPLATE +
    "\n\nВерни ТОЛЬКО полный заполненный текст договора, без пояснений и markdown.";
  const filled = await callClaude(prompt, 8000);
  const lines = filled.split("\n");
  const paras = lines.map(line => {
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
  [{ text: "🎯 Подготовить КП (презентация)", callback_data: "kp" }],
  [{ text: "📄 Подготовить договор", callback_data: "contract" }]
] } };

const REQUISITES_GUIDE =
  "📄 Пришлите <b>реквизиты обеих сторон</b> одним сообщением. Шаблон:\n\n" +
  "<b>ЗАКАЗЧИК (клиент):</b>\n• Наименование / ИП:\n• БИН/ИИН:\n• Юр. адрес:\n• В лице (ФИО, должность):\n• Банковские реквизиты (банк, БИК, счёт IBAN):\n\n" +
  "<b>ИСПОЛНИТЕЛЬ (вы):</b>\n• Наименование / ИП:\n• БИН/ИИН:\n• Юр. адрес:\n• В лице (ФИО, должность):\n• Банковские реквизиты:\n\n" +
  "• Ежемесячная цена (₸):\n\n(Номер договора франшизы подставлю автоматически по исполнителю. Скопируйте шаблон, заполните и отправьте.)";

async function generateKp(chatId, transcript, priceOverride) {
  const wait = await tg("sendMessage", { chat_id: chatId, text: "🎯 Готовлю индивидуальное КП…" });
  try {
    const kp = await extractKp(transcript);
    if (priceOverride) kp.price = priceOverride.replace(/\D/g, "");
    if (!kp.price) {
      state[chatId] = state[chatId] || {}; state[chatId].stage = "await_kp_price"; state[chatId].diagnostic = transcript;
      await send(chatId, "В диагностике не нашёл озвученную цену. Напишите ежемесячную цену, которую вы озвучили клиенту (например: 970000).");
      return;
    }
    const buf = await buildKpPptx(kp, chatId);
    await sendDocument(chatId, "Finpark_KP.pptx", buf, "Индивидуальное КП для: " + (kp.company || "клиента"));
    if (state[chatId]) state[chatId].stage = "idle";
  } catch (e) {
    await send(chatId, "⚠️ Не удалось собрать КП: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

async function generateContract(chatId, requisites) {
  const wait = await tg("sendMessage", { chat_id: chatId, text: "📄 Готовлю договор…" });
  try {
    const price = (state[chatId] && state[chatId].price) || null;
    const buf = await buildContractDocx(requisites, price);
    await sendDocument(chatId, "Finpark_Dogovor.docx", buf, "Договор готов. Проверьте реквизиты, номер франшизы и суммы перед подписанием.");
    if (state[chatId]) state[chatId].stage = "idle";
  } catch (e) {
    await send(chatId, "⚠️ Не удалось собрать договор: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

async function handle(upd) {
  if (upd.callback_query) {
    const cq = upd.callback_query; const chatId = cq.message.chat.id;
    tg("answerCallbackQuery", { callback_query_id: cq.id }).catch(()=>{});
    const st = state[chatId];
    if (cq.data === "kp") {
      if (!st || !st.diagnostic) { await send(chatId, "Сначала пришлите диагностику для оценки."); return; }
      await generateKp(chatId, st.diagnostic);
    } else if (cq.data === "contract") {
      state[chatId] = state[chatId] || {}; state[chatId].stage = "await_contract_requisites";
      await send(chatId, REQUISITES_GUIDE);
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
  if (st && st.stage === "await_contract_requisites" && (text || msg.document)) {
    let req = text;
    if (msg.document) { try { req = await readDocumentText(msg.document) || text; } catch {} }
    if (!req || req.length < 20) { await send(chatId, "Пришлите реквизиты текстом по шаблону."); return; }
    const pm = req.match(/(\d[\d  ]{3,})\s*(?:₸|тенге|тг)/i);
    state[chatId].price = pm ? pm[1].replace(/\D/g, "") : (st.price || null);
    await generateContract(chatId, req); return;
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
    state[chatId] = { stage: "idle", diagnostic: transcript };
    await send(chatId, buildReport(parsed));
    await send(chatId, "Нужно подготовить документы по этому клиенту?", OFFER_KB);
  } catch (e) {
    await send(chatId, "⚠️ Ошибка оценки: " + esc(e.message || String(e)));
  } finally { if (wait.ok) tg("deleteMessage", { chat_id: chatId, message_id: wait.result.message_id }).catch(()=>{}); }
}

async function main() {
  console.log("Finpark bot v3 запущен. Модель:", MODEL, "| код доступа:", ACCESS_CODE ? "включён" : "выключен");
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
