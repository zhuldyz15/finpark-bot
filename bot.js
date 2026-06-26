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
if (!TOKEN) { console.error("Нет BOT_TOKEN"); process.exit(1); }
if (!KEY) { console.error("Нет ANTHROPIC_API_KEY"); process.exit(1); }
const API = "https://api.telegram.org/bot" + TOKEN;
const allowed = new Set();
const state = {}; // chatId -> { stage, diagnostic, price }
const sleep = ms => new Promise(r => setTimeout(r, ms));

const NAVY = "1F3864", BLUE = "2E75B6", GREYT = "595959", LIGHT = "EAF0FA";

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

// ---------- КП: извлечение данных и сборка презентации ----------
async function extractKp(transcript) {
  const prompt =
    "Из ТЕКСТА ДИАГНОСТИКИ извлеки данные для коммерческого предложения Finpark. " +
    "Боли и решения формулируй индивидуально под этого клиента. Цену бери ту, что партнёр озвучил клиенту в тексте (ежемесячный платёж в тенге, только цифры, например 970000); если цена не названа — null.\n\n" +
    "ТЕКСТ:\n" + transcript.slice(0, 60000) +
    '\n\nВерни СТРОГО JSON без markdown: {"company":"название или Клиент","niche":"ниша","pains":["боль1","боль2","боль3","боль4"],"goals":["цель1","цель2"],"solution":["как Finpark закроет боль1","...","..."],"price":"970000" или null}';
  return parseJson(await callClaude(prompt, 1500));
}

function fmtPrice(p) { const n = String(p).replace(/\D/g, ""); return n ? n.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + " ₸" : ""; }

async function buildKpPptx(kp, chatId) {
  const p = new PptxGenJS();
  p.defineLayout({ name: "FP", width: 13.33, height: 7.5 });
  p.layout = "FP";
  const company = kp.company && kp.company !== "Клиент" ? kp.company : "вашей компании";
  const W = 13.33;
  const titleBar = (s, txt) => s.addText(txt, { x: 0.6, y: 0.4, w: W - 1.2, h: 0.9, fontSize: 26, bold: true, color: NAVY });

  let s = p.addSlide(); s.background = { color: NAVY };
  s.addText("FINPARK", { x: 0.6, y: 0.7, w: 6, h: 0.6, fontSize: 20, bold: true, color: "FFFFFF", charSpacing: 3 });
  s.addText("Коммерческое предложение", { x: 0.6, y: 2.5, w: W - 1.2, h: 1, fontSize: 40, bold: true, color: "FFFFFF" });
  s.addText("Финансовая система для бизнеса под ключ", { x: 0.6, y: 3.6, w: W - 1.2, h: 0.7, fontSize: 22, color: "BCD3EE" });
  s.addText("для " + company, { x: 0.6, y: 4.6, w: W - 1.2, h: 0.7, fontSize: 22, italic: true, color: "FFFFFF" });

  s = p.addSlide(); titleBar(s, "О Finpark");
  const stats = [["7+ лет", "на рынке финансового консалтинга"], ["300+", "клиентов по всему миру"], ["93%", "клиентов подписывают второй договор"], ["1+ млрд ₸", "увеличили прибыль клиентов"]];
  stats.forEach((st, i) => {
    const x = 0.6 + (i % 2) * 6.2, y = 1.6 + Math.floor(i / 2) * 2.4;
    s.addText(st[0], { x, y, w: 5.8, h: 0.9, fontSize: 32, bold: true, color: BLUE });
    s.addText(st[1], { x, y: y + 0.9, w: 5.8, h: 0.9, fontSize: 16, color: GREYT });
  });

  s = p.addSlide(); titleBar(s, "Ваша ситуация сейчас");
  s.addText((kp.pains || []).map(t => ({ text: t, options: { bullet: { code: "2022" }, color: "333333", fontSize: 18, paraSpaceAfter: 10 } })), { x: 0.7, y: 1.6, w: W - 1.4, h: 5.2, valign: "top" });

  s = p.addSlide(); titleBar(s, "Решение: финансовый директор на аутсорсе Finpark");
  s.addText((kp.solution || []).map(t => ({ text: t, options: { bullet: { code: "2713" }, color: "333333", fontSize: 18, paraSpaceAfter: 10 } })), { x: 0.7, y: 1.6, w: W - 1.4, h: 5.2, valign: "top" });

  s = p.addSlide(); titleBar(s, "Что входит: 6 месяцев сопровождения");
  const incl = ["Управленческая отчётность: ОПиУ, ДДС, Балансовый отчёт + дашборды", "Финансовая модель на 12 месяцев", "6 стратегических сессий с менеджментом", "До 2 онлайн-встреч с финдиректором в неделю", "24 планёрки по контролю расходов", "Защита цифр перед инвесторами, дорожная карта проекта", "Чат с финдиректором в рабочее время"];
  s.addText(incl.map(t => ({ text: t, options: { bullet: { code: "2022" }, color: "333333", fontSize: 16, paraSpaceAfter: 8 } })), { x: 0.7, y: 1.6, w: W - 1.4, h: 5.4, valign: "top" });

  s = p.addSlide(); titleBar(s, "Финдиректор Finpark vs штатный");
  const rows = [[{ text: "Критерий", options: { bold: true, color: "FFFFFF", fill: NAVY } }, { text: "Finpark", options: { bold: true, color: "FFFFFF", fill: NAVY } }, { text: "Штатный", options: { bold: true, color: "FFFFFF", fill: NAVY } }],
    ["Оплата", "За фактически оказанные услуги", "Полный оклад + налоги + льготы"],
    ["Стоимость в месяц", "от 400–600 тыс ₸", "от 1 705 тыс ₸ (с налогами)"],
    ["Команда", "ФД + РОК + менеджер + ОКК", "Один специалист"],
    ["Ответственность", "За финансовый результат", "На собственнике"]];
  s.addTable(rows, { x: 0.6, y: 1.6, w: W - 1.2, colW: [3, 4.5, 4.6], fontSize: 14, border: { type: "solid", color: "DDDDDD" }, valign: "middle", rowH: 0.7 });

  s = p.addSlide(); titleBar(s, "Результат работы с Finpark");
  const res = ["Порядок в финансах: внедрены финмодель и управленческие отчёты", "Контроль расходов: устранены кассовые разрывы, календарь платежей", "Устойчивая генерация прибыли 3 месяца подряд", "Регулярные дивиденды без ущерба для компании", "Данные для управленческих решений на основе твёрдых цифр"];
  s.addText(res.map(t => ({ text: t, options: { bullet: { code: "2713" }, color: "333333", fontSize: 18, paraSpaceAfter: 10 } })), { x: 0.7, y: 1.6, w: W - 1.4, h: 5.2, valign: "top" });

  s = p.addSlide(); s.background = { color: LIGHT }; titleBar(s, "Стоимость");
  const price = kp.price ? fmtPrice(kp.price) : null;
  if (price) {
    s.addText(price + " / месяц", { x: 0.6, y: 2.0, w: W - 1.2, h: 1.2, fontSize: 44, bold: true, color: NAVY });
    s.addText("Минимальный срок работы по договору — 6 месяцев.", { x: 0.6, y: 3.3, w: W - 1.2, h: 0.6, fontSize: 18, color: GREYT });
  } else {
    s.addText("Индивидуальные условия обсуждаются с вашим финансовым директором.", { x: 0.6, y: 2.2, w: W - 1.2, h: 1.2, fontSize: 22, color: NAVY });
  }
  s.addText("Окупаемость услуг — за счёт найденных и высвобожденных внутри бизнеса средств.", { x: 0.6, y: 4.4, w: W - 1.2, h: 0.8, fontSize: 16, italic: true, color: GREYT });

  s = p.addSlide(); s.background = { color: NAVY };
  s.addText("Увеличьте финансовые показатели уже сейчас", { x: 0.6, y: 2.6, w: W - 1.2, h: 1.2, fontSize: 30, bold: true, color: "FFFFFF" });
  s.addText("Свяжитесь с вашим финансовым директором Finpark для старта.", { x: 0.6, y: 3.9, w: W - 1.2, h: 0.8, fontSize: 18, color: "BCD3EE" });

  const file = path.join("/tmp", "kp_" + chatId + "_" + Date.now() + ".pptx");
  await p.writeFile({ fileName: file });
  const buf = fs.readFileSync(file); fs.unlink(file, () => {});
  return buf;
}

// ---------- Договор ----------
async function buildContractDocx(requisites, price) {
  const prompt =
    "Заполни ПРОПУСКИ в договоре (отмечены подчёркиваниями) данными из РЕКВИЗИТОВ. " +
    "Стоимость услуг за 6 месяцев и ежемесячный платёж рассчитай из ежемесячной цены: ежемесячно = " + (price || "указанной в реквизитах") + " тенге, за 6 месяцев = ежемесячно × 6. " +
    "Везде, где требуется «(сумма прописью)», впиши сумму словами на русском. Дату договора поставь сегодняшней. " +
    "Сохрани ВЕСЬ текст и структуру договора без изменений, только заполни пропуски. Если каких-то данных нет — оставь короткий пропуск «____».\n\n" +
    "РЕКВИЗИТЫ:\n" + requisites + "\n\nШАБЛОН ДОГОВОРА:\n" + CONTRACT_TEMPLATE +
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
  "• Ежемесячная цена (₸):\n\n(Скопируйте шаблон, заполните и отправьте.)";

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
    await sendDocument(chatId, "Finpark_Dogovor.docx", buf, "Договор готов. Проверьте реквизиты и суммы перед подписанием.");
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
  console.log("Finpark bot v2 запущен. Модель:", MODEL, "| код доступа:", ACCESS_CODE ? "включён" : "выключен");
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
