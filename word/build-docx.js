/**
 * build-docx.js
 * יוצר 4 קבצי DOCX מקצועיים עם RTL מלא למינהל גמלאות, ביטוח לאומי
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, BorderStyle, WidthType, ShadingType,
} = require('docx');
const fs = require('fs');

// ─── BiDi helpers ────────────────────────────────────────────────────────────
const LRE = '\u202A';
const PDF = '\u202C';
const RLM = '\u200F';
function eng(text) { return `${LRE}${text}${PDF}${RLM}`; }

// ─── Base factories ───────────────────────────────────────────────────────────
function r(text, { bold=false, size=19, color='1A1A1A', font='Arial', italic=false }={}) {
  return new TextRun({ text, bold, size, color, font, rtl: true, italics: italic });
}

function p(children, { spacing={ before:0, after:50 }, numbering, indent, border, align=AlignmentType.RIGHT }={}) {
  return new Paragraph({
    children,
    bidirectional: true,
    alignment: align,
    spacing,
    ...(numbering ? { numbering } : {}),
    ...(indent    ? { indent }    : {}),
    ...(border    ? { border }    : {}),
  });
}

function spacer(pts=80) {
  return p([], { spacing:{ before:pts, after:0 } });
}

const noBorders = {
  top:     { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
  bottom:  { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
  left:    { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
  right:   { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
  insideH: { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
  insideV: { style: BorderStyle.NONE, size:0, color:'FFFFFF' },
};

const RTL_NUMBERING_CONFIG = [{
  reference: 'rtlBullets',
  levels: [{ level:0, format:LevelFormat.BULLET, text:'\u25AA', alignment:AlignmentType.RIGHT,
    style:{ paragraph:{ indent:{ right:400, hanging:400 }, bidirectional:true } } }],
}];

function bullet(text, size=18) {
  return p([r(text, { size })], {
    numbering: { reference:'rtlBullets', level:0 },
    spacing: { before:0, after:35 },
  });
}

// ─── UI Components ────────────────────────────────────────────────────────────

/** שורת כותרת עליונה — לוגו + תאריך */
function topBar() {
  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: [4600, 4600],
    borders: noBorders,
    rows: [new TableRow({ children:[
      new TableCell({
        borders: noBorders,
        shading: { fill:'1B3A6B', type:ShadingType.CLEAR },
        margins: { top:80, bottom:80, left:200, right:200 },
        width: { size:4600, type:WidthType.DXA },
        children: [p([r(`ב"ל | ביטוח לאומי`, { bold:true, size:17, color:'FFFFFF' })], { spacing:{ before:0, after:0 } })],
      }),
      new TableCell({
        borders: noBorders,
        shading: { fill:'1B3A6B', type:ShadingType.CLEAR },
        margins: { top:80, bottom:80, left:200, right:200 },
        width: { size:4600, type:WidthType.DXA },
        children: [p([r('מינהל גמלאות | מרץ 2026', { size:17, color:'9DC3E6' })],
          { align:AlignmentType.LEFT, spacing:{ before:0, after:0 } })],
      }),
    ]})]
  });
}

/** בלוק כותרת ראשי עם רקע כהה */
function titleBlock(icon, title, subtitle, statusText, statusColor='2E7D32') {
  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: [9200],
    borders: noBorders,
    rows: [new TableRow({ children:[
      new TableCell({
        borders: noBorders,
        shading: { fill:'1F3864', type:ShadingType.CLEAR },
        margins: { top:140, bottom:110, left:220, right:220 },
        width: { size:9200, type:WidthType.DXA },
        children: [
          p([r(`${icon}  ${title}`, { bold:true, size:28, color:'FFFFFF' })], { spacing:{ before:0, after:40 } }),
          p([r(subtitle, { size:18, color:'BDD7EE' })], { spacing:{ before:0, after:50 } }),
          p([
            r('סטטוס: ', { size:16, color:'9DC3E6' }),
            r(statusText, { size:16, color:'A8E6CF', bold:true }),
          ], { spacing:{ before:0, after:0 } }),
        ],
      }),
    ]})]
  });
}

/** כותרת סקציה עם רקע כחול */
function sectionHeader(label) {
  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: [9200],
    borders: noBorders,
    rows: [new TableRow({ children:[
      new TableCell({
        borders: noBorders,
        shading: { fill:'2E5F8A', type:ShadingType.CLEAR },
        margins: { top:70, bottom:70, left:180, right:180 },
        width: { size:9200, type:WidthType.DXA },
        children: [p([r(label, { bold:true, size:19, color:'FFFFFF' })], { spacing:{ before:0, after:0 } })],
      }),
    ]})]
  });
}

/** תיבת תוכן עם רקע בהיר */
function contentBox(children, bg='EBF3FB') {
  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: [9200],
    borders: noBorders,
    rows: [new TableRow({ children:[
      new TableCell({
        borders: {
          ...noBorders,
          right: { style:BorderStyle.SINGLE, size:12, color:'2E5F8A' },
        },
        shading: { fill:bg, type:ShadingType.CLEAR },
        margins: { top:100, bottom:100, left:200, right:200 },
        width: { size:9200, type:WidthType.DXA },
        children,
      }),
    ]})]
  });
}

/** 3 תיבות KPI בשורה אחת */
function kpiRow(kpis) {
  // kpis = [{num, label, sub?}, ...]  — exactly 3
  const colW = 2900;
  const gap  = 250;
  const cols = [colW, gap, colW, gap, colW];

  const cells = kpis.flatMap((k, i) => {
    const dataCell = new TableCell({
      borders: { top:{ style:BorderStyle.SINGLE, size:8, color:'2E5F8A' }, ...noBorders },
      shading: { fill:'DDEEFF', type:ShadingType.CLEAR },
      margins: { top:110, bottom:110, left:140, right:140 },
      width: { size:colW, type:WidthType.DXA },
      children: [
        p([r(k.num, { bold:true, size:32, color:'1F3864' })], { align:AlignmentType.CENTER, spacing:{ before:0, after:20 } }),
        p([r(k.label, { size:17, color:'2E5F8A', bold:true })], { align:AlignmentType.CENTER, spacing:{ before:0, after:k.sub?14:0 } }),
        ...(k.sub ? [p([r(k.sub, { size:15, color:'555555' })], { align:AlignmentType.CENTER, spacing:{ before:0, after:0 } })] : []),
      ],
    });
    if (i < 2) {
      const gapCell = new TableCell({
        borders: noBorders,
        shading: { fill:'FFFFFF', type:ShadingType.CLEAR },
        width: { size:gap, type:WidthType.DXA },
        children: [p([r('')])],
      });
      return [dataCell, gapCell];
    }
    return [dataCell];
  });

  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: cols,
    borders: noBorders,
    rows: [new TableRow({ children: cells })],
  });
}

/** תיבת CTA / החלטה נדרשת */
function ctaBox(children) {
  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: [9200],
    borders: noBorders,
    rows: [new TableRow({ children:[
      new TableCell({
        borders: {
          ...noBorders,
          bottom: { style:BorderStyle.SINGLE, size:12, color:'1B3A6B' },
          top:    { style:BorderStyle.SINGLE, size:12, color:'1B3A6B' },
        },
        shading: { fill:'FFF9C4', type:ShadingType.CLEAR },
        margins: { top:110, bottom:110, left:200, right:200 },
        width: { size:9200, type:WidthType.DXA },
        children,
      }),
    ]})]
  });
}

/** שורת קרדיט + קישור */
function footerBar(link='') {
  return new Table({
    width: { size:9200, type:WidthType.DXA },
    columnWidths: [5200, 4000],
    borders: noBorders,
    rows: [new TableRow({ children:[
      new TableCell({
        borders: { ...noBorders, top:{ style:BorderStyle.SINGLE, size:6, color:'BBBBBB' } },
        shading: { fill:'F5F5F5', type:ShadingType.CLEAR },
        margins: { top:80, bottom:80, left:200, right:200 },
        width: { size:5200, type:WidthType.DXA },
        children: [p([r('אביעד יצחקי | מינהל גמלאות | ביטוח לאומי', { size:15, color:'555555' })],
          { spacing:{ before:0, after:0 } })],
      }),
      new TableCell({
        borders: { ...noBorders, top:{ style:BorderStyle.SINGLE, size:6, color:'BBBBBB' } },
        shading: { fill:'F5F5F5', type:ShadingType.CLEAR },
        margins: { top:80, bottom:80, left:200, right:200 },
        width: { size:4000, type:WidthType.DXA },
        children: [p(link ? [r(link, { size:15, color:'2E5F8A', italic:true })] : [r('', { size:15 })],
          { align:AlignmentType.LEFT, spacing:{ before:0, after:0 } })],
      }),
    ]})]
  });
}

// ─── Document builder helper ──────────────────────────────────────────────────
function buildDoc(children) {
  return new Document({
    numbering: { config: RTL_NUMBERING_CONFIG },
    styles: {
      default: {
        document: { run: { font:'Arial', size:19, rtl:true } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width:11906, height:16838 },
          margin: { top:650, right:750, bottom:650, left:750 },
        }
      },
      children,
    }]
  });
}

async function save(doc, path) {
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(path, buf);
  console.log(`✓ ${path}`);
}

const OUT = 'C:\\Users\\btl\\Desktop\\פרוייקטים\\one-pagers\\word\\';

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 1: דשבורד פערי גמלאות
// ═══════════════════════════════════════════════════════════════════════════════
async function buildDashboard() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '📊',
      'דשבורד פערי גמלאות ברשויות',
      'מפת הפערים הארצית — לראשונה, בזמן אמת',
      '✅ פעיל — כלי עבודה בשטח'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('החלטות על חלוקת משאבים התקבלו ללא תמונת מצב ארצית — אין נראות לפערי מיצוי בין רשויות, אין כלי השוואה, אין דרך לזהות אוכלוסיות שנופלות בין הכיסאות.', { size:18 })]),
      spacer(16),
      p([r(`280 רשויות, 9 גמלאות, 339 ישובים — עד היום לא הייתה תמונה אחת שמרכזת את כולם.`, { size:18 })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('דשבורד אנליטי אינטראקטיבי מבוסס נתוני אמת מביטוח לאומי (דצמבר 2024), מנורמלים לפי אוכלוסיית יעד לכל רשות.', { size:18 })]),
      spacer(14),
      bullet('שלושה מסכים: סקירה כללית, ניתוח לפי גמלה, השוואת רשויות'),
      bullet('סינון לפי אשכול, מחוז, גמלה — מפה אינטראקטיבית שחושפת פערים בשנייה'),
      bullet('9 גמלאות: נכות, סיעוד, ילד נכה, ניידות, אבטלה, הבטחת הכנסה, מזונות, זקנה, ילדים'),
      spacer(14),
      p([r('הכלי יצא לשטח ומחולל תובנות — פערים משמעותיים עולים בכלל הגמלאות ובכלל הרשויות.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'280', label:'רשויות מקומיות', sub:'מוצגות בדשבורד' },
      { num:'9',   label:'סוגי גמלאות',    sub:'מנורמלים לאוכלוסייה' },
      { num:'339', label:'ישובים',         sub:'במאגר הנתונים' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('הדשבורד מתעדכן ומשוכלל: דיוק חישוב זקנה, הוספת 55 מועצות אזוריות, התאמה ל-24 סניפים.', { size:18 })]),
      spacer(14),
      p([
        r('החלטה נדרשת: ', { bold:true, size:18, color:'1B3A6B' }),
        r('אימוץ הדשבורד ככלי קבוע בישיבות הנהלה ובדיוני הקצאת משאבים מחוזיים.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar('https://avi1840.github.io/dashboard-gamlaut/'),
  ]);
  await save(doc, OUT + '03-dashboard-gamlaut.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 2: סל אישי בסיעוד
// ═══════════════════════════════════════════════════════════════════════════════
async function buildSalIshi() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🏥',
      'סל אישי בסיעוד — רפורמת הזדקנות מיטבית',
      'מעבר ממודל סיעוד אחיד לסל שירותים מותאם אישית',
      `🟡 ${eng('Osprint')} בתכנון — מתחיל אחרי פסח 2026`
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('220,000 זכאי סיעוד מקבלים סל אחיד — שעות טיפול אישי — ללא התחשבות בצרכים, תרבות או מיקום.', { size:18 })]),
      spacer(14),
      bullet(`44.6% מזכאי רמה 1 שהידרדרו לא קיבלו אף שירות נוסף מעבר לטיפול האישי הבסיסי (${eng('SDI')}=0)`),
      bullet(`נשים בנות 85+ שוהות 40.3 חודשים ברמה 1 לפני הידרדרות — ${eng('RDI')} של 1.49 לעומת נורמה של 27 חודשים`),
      spacer(14),
      p([r('המערכת מזהה הידרדרות רק כשהיא כבר קרתה — אין מנגנון מניעה.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('רפורמת סל אישי: מתאמת שירות מעריכה ב-6 ממדים ובונה סל מותאם מתוך קטלוג שירותים קהילתיים — לפי אזור, תרבות ורמת תפקוד.', { size:18 })]),
      spacer(14),
      bullet(`מנוע מדדי הידרדרות (${eng('RDI')} + ${eng('SDI')}) לזיהוי מוקדם וטיוב ההקצאה לפני החמרה`),
      bullet(`אבטיפוס עובד (${eng('Lovable')}): דשבורד מתאמת, קליטה, הערכה, המלצת סל, מעקב הידרדרות`),
      bullet('הכלי בנוי גנרי לשימוש רוחבי: נכויות, שיקום, אבטלה'),
      spacer(14),
      p([r('פיילוט ירושלים "עכשיו אני" — 286 משתתפים ברמות 1-3 — קרקע הפעלה ראשונה מוכנה.', { size:18 })]),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'220K',  label:'זכאי סיעוד',        sub:'אוכלוסיית יעד' },
      { num:'44.6%', label:'הידרדרו ללא שירות', sub:'נוסף (SDI=0, רמה 1)' },
      { num:'286',   label:'משתתפי פיילוט',     sub:'ירושלים' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r(`${eng('Osprint')} עם שטראוס אסטרטגיה מתחיל אחרי פסח 2026 — צוות ליבה: ביטוח לאומי + שטראוס + ירושלים.`, { size:18 })]),
      spacer(14),
      p([r('מפגש פתיחה: הצגת פיילוט ירושלים + הדגמת האבטיפוס העובד.', { size:18 })]),
      spacer(10),
      p([
        r('החלטה נדרשת: ', { bold:true, size:18, color:'1B3A6B' }),
        r('אישור הרכב הצוות ותאריכי ה-Osprint.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar(),
  ]);
  await save(doc, OUT + '11-sal-ishi.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 3: פורטל גמלאות 2.0
// ═══════════════════════════════════════════════════════════════════════════════
async function buildPortal() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🗂️',
      'פורטל גמלאות 2.0 — מגדלור הידע המקצועי',
      'בוטים על חוזרים, תדריכים ופסקי דין — הידע נגיש לפקיד בשניות',
      '🔄 שלב ב — בוטים בעבודה, עיצוב כמעט מוכן'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('4,500 עובדי גמלאות מחפשים מידע מקצועי ב-123 חוזרים, תדריכים ופסקי דין — ללא חיפוש חכם, ללא בינה מלאכותית, ללא הנגשה מובנית.', { size:18 })]),
      spacer(14),
      p([r('הידע קיים בפורטל — אבל לא נגיש: עובד שמחפש הוראה ספציפית בחוזר מבלה דקות יקרות בחיפוש ידני, לעיתים ללא תוצאה.', { size:18 })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('שדרוג בשני מסלולים במקביל:', { size:18, bold:true })]),
      spacer(14),
      bullet('מסלול ראשי — בוטים על הידע המקצועי: סקריפט שיוצא תוכן מהפורטל הקיים ומזינה בוטים AI שינגישו חוזרים, תדריכים ופסקי דין בשיחה טבעית'),
      bullet('מסלול משני — עיצוב חדש: דף הבית, נפגעי עבודה והבטחת הכנסה עוברים עיצוב ראשוני — כמעט מוכן, עם שיתוף פעולה פעיל מול מחשוב'),
      spacer(14),
      p([r('זהו השדרוג שישנה את חוויית העבודה היומיומית של 4,500 עובדים.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'4,500', label:'עובדי גמלאות',     sub:'יפיקו תועלת' },
      { num:'123',   label:'חוזרים ותדריכים', sub:'יונגשו בבוטים' },
      { num:'דקות→שניות', label:'זמן חיפוש', sub:'מקצועי' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('הסקריפט לחילוץ תוכן הפורטל בעבודה — כדי להעלות את הבוטים לאוויר נדרשים: שרת + גישה ל-API.', { size:18 })]),
      spacer(14),
      p([
        r('החלטות נדרשות: ', { bold:true, size:18, color:'1B3A6B' }),
      ]),
      bullet('הקצאת שרת ו-API לסביבת הבוטים — בידי מחשוב'),
      bullet('קיצור הזמן בין השלמת העיצוב לביצוע בפועל'),
    ]),
    spacer(20),
    footerBar(),
  ]);
  await save(doc, OUT + '07-portal-gamlaut.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 4: בוט סיכום תביעת סיעוד
// ═══════════════════════════════════════════════════════════════════════════════
async function buildBotSiud() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🤖',
      'בוט סיכום תביעת סיעוד',
      `חיסכון של ~6 שעות ביום למעריך — ${eng('MVP')} מוכן, מחכה לאות ירוק`,
      `✅ ${eng('MVP')} מוכן — תכנית עבודה קיימת, לא מתקדמת בדיגיטל`
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('מעריכי סיעוד מטפלים ב-8 עד 12 תביעות ביום — כל סיכום לוקח 30 עד 45 דקות ידנית: תיעוד, שאלות תפקוד, הכנת מסמכים לוועדה.', { size:18 })]),
      spacer(14),
      p([r('עבודה שגרתית, חוזרת ולא-אחידה — תוצאה: עומס, פערי איכות בין מעריכים, וזמן יקר שלא מופנה לשיקול מקצועי.', { size:18 })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r(`בוט ${eng('AI')} (גרסת פרומפט ${eng('v4+')}): מקבל נתוני תביעה ומייצר בשניות — סיכום מובנה ואחיד, שאלות תפקוד מותאמות למצב הרפואי, הכנת מסמכים לוועדה.`, { size:18 })]),
      spacer(14),
      bullet('שיקול הדעת המקצועי של המעריך נשמר — הבוט הוא כלי עזר, לא תחליף'),
      bullet('כולל סימולציות קבוצתיות לתרגול צוותי'),
      bullet(`~30 מסמכי פיתוח ואפיון, תכנית עבודה מוכנה ומפורטת`),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'~6 שעות', label:'חיסכון ביום',     sub:'למעריך (10×37 דקות)' },
      { num:'8–12',    label:'תביעות ביום',      sub:'למעריך ממוצע' },
      { num:`${eng('v4+')}`,      label:'גרסת פרומפט',    sub:'מפותחת ומוכנה' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('תכנית עבודה מוכנה ומפורטת — אך לא מתקדמת בדיגיטל.', { size:18 })]),
      spacer(14),
      p([
        r('החלטה נדרשת: ', { bold:true, size:18, color:'1B3A6B' }),
        r('אישור המשך מול הדיגיטל — בחירת פלטפורמה והפעלת פיילוט עם 5 עד 10 מעריכי סיעוד בסניף אחד.', { size:18 }),
      ]),
      spacer(14),
      p([r('כל שבוע של עיכוב = ~30 שעות עבודה לא-מנוצלות לכל מעריך.', { size:18, italic:true, color:'8B0000' })]),
    ]),
    spacer(20),
    footerBar(),
  ]);
  await save(doc, OUT + '06-bot-siud.docx');
}

// ─── Run all ──────────────────────────────────────────────────────────────────
(async () => {
  console.log('\nבונה מסמכי DOCX...\n');
  await buildDashboard();
  await buildSalIshi();
  await buildPortal();
  await buildBotSiud();
  console.log('\n✅ 4 מסמכים נוצרו בהצלחה ב:', OUT);
  console.log('\nהרץ כעת:');
  console.log(`python3 post-process.py <file>.docx <file>.docx`);
})();
