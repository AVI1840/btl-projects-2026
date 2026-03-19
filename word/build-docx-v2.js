/**
 * build-docx-v2.js
 * יוצר 7 קבצי DOCX נוספים — ארנק, מחשבון, מיצוי 360, ועדות, סיעוד, בחירת גמלה, עוזר אישי
 * אותה ארכיטקטורה כמו build-docx.js
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

function kpiRow(kpis) {
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
// DOCUMENT 1: ארנק זכויות
// ═══════════════════════════════════════════════════════════════════════════════
async function buildArnak() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '👛',
      'ארנק זכויות — מצפן ההטבות',
      'רשימת זכויות אישית מותאמת לכל אזרח — לפי פרופיל חיים',
      '✅ חי באוויר — 98/98 הטבות מלאות'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('אזרחים לא יודעים מה מגיע להם. המידע על 98 הטבות ביטוח לאומי מפוזר בין אתרים, חוזרים ונהלים. אין "מקום אחד" שכל אזרח יכול לפנות אליו ולקבל תמונה שלמה של מה שמגיע לו — לפי הפרופיל האישי שלו.', { size:18 })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('ממשק אזרח ידידותי שמנחה בשאלות פשוטות ומייצר "ארנק זכויות" אישי עם כל ההטבות הרלוונטיות. מבוסס React + Framer Motion, פרוס ב-GitHub Pages.', { size:18 })]),
      spacer(14),
      bullet('98 הטבות ב-8 קטגוריות — מיין לפי 13 סוגי גמלה'),
      bullet('ממשק RTL מלא, עיצוב ממשלתי, נגישות WCAG'),
      bullet('פועל ללא שרת — static site, אפס עלות תפעולית'),
      bullet('FeedbackModal מובנה עם Google Sheets לאיסוף משוב'),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'98',  label:'הטבות במאגר',   sub:'8 קטגוריות' },
      { num:'13',  label:'סוגי גמלה',      sub:'מסווגים ומסוננים' },
      { num:'0₪',  label:'עלות תפעולית',  sub:'GitHub Pages חינמי' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('הכלי חי ומוכן לשימוש — מחכה לפיילוט רשמי עם פקידים ואזרחים.', { size:18 })]),
      spacer(14),
      p([
        r('החלטה נדרשת: ', { bold:true, size:18, color:'1B3A6B' }),
        r('אישור לפיילוט מבוקר עם 10–20 פקידי שירות בסניף אחד — מדידת שיפור במיצוי זכויות לפני ואחרי.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar('https://aviad1840.github.io/arnak-zchuyot/'),
  ]);
  await save(doc, OUT + '01-arnak-zchuyot.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 2: מחשבון כפל גמלאות
// ═══════════════════════════════════════════════════════════════════════════════
async function buildMachshevon() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🧮',
      'מחשבון כפל גמלאות — סעיף 320',
      'חישוב אוטומטי ומיידי של זכאות לכפל גמלאות לפי חוק הביטוח הלאומי',
      '✅ חי באוויר — בפיילוט 2026'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('סעיף 320 לחוק הביטוח הלאומי קובע מגבלות על כפל גמלאות — אך הוא מורכב ורב-שכבתי. פקידים מתקשים לחשב ידנית תוך כדי שיחה. אזרחים רבים לא יודעים שהם זכאים, וטעויות חישוב עולות ביוקר לשני הצדדים.', { size:18 })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('מחשבון אינטראקטיבי שמקבל את סוגי הגמלאות של האזרח ומחשב באופן מיידי את הזכאות לכפל — כולל הסבר מפורט לכל תוצאה.', { size:18 })]),
      spacer(14),
      bullet('כיסוי מלא של כל מקרי סעיף 320'),
      bullet('כלי לשימוש הפקיד בזמן השיחה — ולאזרח לבדיקה עצמית'),
      bullet('תוצאה תוך שניות עם נימוק ברור לכל החלטה'),
      spacer(14),
      p([r('⚖️ נדרש אישור משפטי לפני הטמעה מבצעית — לוגיקה מוכנה לבדיקה.', { size:18, bold:true, color:'8B5A00' })]),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'320',     label:'סעיף בחוק',      sub:'ביטוח לאומי' },
      { num:'שניות',  label:'זמן חישוב',       sub:'תוצאה מיידית' },
      { num:'100%',   label:'כיסוי מקרים',     sub:'כל תרחישי סעיף 320' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('אימות מול יועצים משפטיים — בדיקה שהלוגיקה תואמת את הפרשנות המשפטית העדכנית.', { size:18 })]),
      spacer(14),
      p([
        r('החלטה נדרשת: ', { bold:true, size:18, color:'1B3A6B' }),
        r('אחרי אישור משפטי — פיילוט עם פקידי גמלאות נבחרים ומדידת שיפור בדיוק החישובים.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar('https://aviad1840.github.io/machshevon-kefel/'),
  ]);
  await save(doc, OUT + '02-machshevon-kefel.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 3: מיצוי זכויות 360°
// ═══════════════════════════════════════════════════════════════════════════════
async function buildMitzui() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🎯',
      'מיצוי זכויות 360°',
      'כלי AI פרואקטיבי לזיהוי זכויות לא ממומשות — תסריט שיחה לפקיד תוך 30 שניות',
      '✅ חי באוויר — פיילוט 2026'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('מיליארדי שקלים בזכויות לא ממומשות בכל שנה — אזרחים לא יודעים מה מגיע להם, ופקידי שירות לא יכולים לבדוק 98 הטבות שונות בשיחה של 5 דקות.', { size:18 })]),
      spacer(14),
      p([r('הפער הזה מתרחב, ופוגע דווקא במי שהכי זקוק לסיוע.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('מערכת AI שמקבלת פרופיל אזרח ומייצרת תוך 30 שניות: רשימת זכויות מותאמת + תסריט שיחה לפקיד.', { size:18 })]),
      spacer(14),
      bullet('7 תרחישי אזרח: ותיק, הורה לילד עם מוגבלות, אלמנה, חד-הורי, עולה חדש, נפגע עבודה, מקבל הבטחת הכנסה'),
      bullet('98 הטבות במאגר הידע — כל אחת עם תנאי זכאות מפורטים'),
      bullet('פלט מובנה: מה מגיע, למה, ואיך לפתוח'),
      spacer(14),
      p([r('הכלי פועל עם Claude API — נדרש רישיון ארגוני לפרויקציה מלאה.', { size:18, color:'8B5A00' })]),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'7',   label:'תרחישי אזרח',    sub:'מכוסים מלאים' },
      { num:'98',  label:'הטבות במאגר',    sub:'ידע עדכני' },
      { num:`30"`, label:'זמן לתוצאה',     sub:'תסריט שיחה מלא' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('פיילוט עם 10 פקידי שירות בסניף אחד — מדידת שיפור במיצוי זכויות לפני ואחרי.', { size:18 })]),
      spacer(10),
      bullet('אישור הנהלה לפיילוט'),
      bullet('סניף מגויס ומוכן'),
      bullet('מדדי הצלחה מוסכמים'),
      spacer(10),
      p([
        r('עלות ריצה: ', { bold:true, size:18, color:'1B3A6B' }),
        r('~200$ לחודש לפיילוט (Claude API) — החזר ROI מיידי.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar('https://aviad1840.github.io/mitzui-360/'),
  ]);
  await save(doc, OUT + '04-mitzui-360.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 4: ועדות רפואיות
// ═══════════════════════════════════════════════════════════════════════════════
async function buildVaadot() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🏥',
      'ועדות רפואיות — צ\'קליסט מסמכים חכם',
      'הכנה מושלמת לוועדה רפואית — המסמך הנכון, בזמן הנכון',
      '✅ חי באוויר — פרוס ב-GitHub Pages, מרץ 2026'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('אזרחים מגיעים לוועדות רפואיות ללא המסמכים הנדרשים. זה גורם לדחיות, עיכובים ארוכים, תסכול — ועלויות מיותרות למערכת.', { size:18 })]),
      spacer(14),
      p([r('ועדה שנדחית מעלה עלויות, מאריכה תהליכים, ופוגעת בזכאות שלא ממומשת.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('צ\'קליסט אינטראקטיבי המותאם לסוג הוועדה הספציפי. האזרח מסמן מה יש לו, המערכת מציגה מה חסר ומסבירה למה כל מסמך חשוב.', { size:18 })]),
      spacer(14),
      bullet('6 סוגי ועדות: נכות כללית, נפגעי עבודה, נכות ילד, סיעוד, אובדן כושר, שירותים מיוחדים'),
      bullet('כלי לשימוש האזרח לפני הוועדה — וגם לפקיד כמדריך הכנה'),
      bullet('ממשק React מלא, RTL, נגישות מובנית'),
      bullet('פרוס כ-static site — אפס עלות תפעולית'),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'6',    label:'סוגי ועדות',      sub:'מכוסים' },
      { num:'0₪',  label:'עלות תפעולית',    sub:'GitHub Pages' },
      { num:'2שע', label:'זמן להטמעה',      sub:'מוכן לפרודקשן' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('הכלי חי ומוכן — מחכה לפיילוט עם 50 אזרחים לפני וועדה.', { size:18 })]),
      spacer(14),
      bullet('פיילוט מבוקר: 50 אזרחים לפני ועדה — מדידת שיפור בהכנה'),
      bullet('שלב ב: שליחת קישור אוטומטית עם הזמנה לוועדה'),
      spacer(10),
      p([
        r('הזדמנות מיידית: ', { bold:true, size:18, color:'1B3A6B' }),
        r('שלח לאזרחים את הקישור עם ההזמנה לוועדה — הכנה אוטומטית, אפס עלות.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar('https://aviad1840.github.io/vaadot-refuiyot/'),
  ]);
  await save(doc, OUT + '05-vaadot-refuiyot.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 5: הנגשת שירותי סיעוד
// ═══════════════════════════════════════════════════════════════════════════════
async function buildNursing() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '♿',
      'הנגשת בחירה ומיצוי זכויות בסיעוד',
      'ממשק React נגיש לסיוע לאזרח בהבנת זכאות סיעוד ומיצוי מלא של זכויותיו',
      '◔ בפיתוח — אפיון מלא, ממשק עובד'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('מתוך 220,000 מקבלי גמלת סיעוד, רבים אינם ממצים את מלוא זכויותיהם — בשל חוסר ידע, בעיות שפה, גיל מתקדם או מוגבלות.', { size:18 })]),
      spacer(14),
      p([r('הפוטנציאל הכלכלי הלא ממומש מגיע לעשרות מיליוני שקלים בשנה.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('ממשק React נגיש ופשוט שמנחה את המבוטח (או בן משפחה) שלב אחר שלב:', { size:18 })]),
      spacer(14),
      bullet('בדיקת זכאות — שאלות פשוטות, תוצאה ברורה'),
      bullet('בחירת מסלול מטפל/שירות — עם הסבר מלא'),
      bullet('הצגת כלל ההטבות המגיעות — ללא השמטות'),
      bullet('נגישות מלאה לפי תקן 5568 — תמיכה בקורא מסך, ניגודיות גבוהה'),
      bullet('תמיכה בריבוי שפות (עברית/ערבית) — לאוכלוסיות מגוונות'),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'220K',  label:'מקבלי סיעוד',     sub:'אוכלוסיית יעד' },
      { num:'5568',  label:'תקן נגישות',       sub:'תאימות מלאה' },
      { num:'React', label:'טכנולוגיה',        sub:'מודולרי ומותאם' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('השלמת פיתוח ובדיקות שמישות עם קבוצת מיקוד של מבוטחים.', { size:18 })]),
      spacer(10),
      bullet('אישור להשלמת פיתוח'),
      bullet('גיוס 15–20 מבוטחים לבדיקות שמישות'),
      bullet('תיאום עם מחלקת IT לאינטגרציה עתידית'),
      bullet('הגדרת מדדי מיצוי לפני ואחרי הטמעה'),
    ]),
    spacer(20),
    footerBar('https://aviad1840.github.io/nursing-helper/'),
  ]);
  await save(doc, OUT + '08-nursing-helper.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 6: בחירת גמלה
// ═══════════════════════════════════════════════════════════════════════════════
async function buildBchirat() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '⚖️',
      'בחירת גמלה — כפל נכות ושאירים',
      'כלי אינטראקטיבי לסיוע לאזרחים הזכאים לשתי גמלאות בקבלת החלטה מושכלת',
      '● אפיון מלא + שלד UI — מוכן לפיתוח'
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('כ-423 אזרחים מדי שנה עומדים בפני בחירה מורכבת בין גמלת נכות לגמלת שאירים — עם השלכות כלכליות ארוכות טווח.', { size:18 })]),
      spacer(14),
      p([r('רובם לא מבינים את ההבדלים ומקבלים החלטה מוטעית — שאינה ניתנת לשינוי.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r('ממשק UI חכם שמקבל נתוני האזרח, מחשב ומשווה את שתי האפשרויות — ומוביל לקבלת ההחלטה האטרקטיבית ביותר.', { size:18 })]),
      spacer(14),
      bullet('חישוב מלא: גמלת נכות מול גמלת שאירים — כולל השלכות מס וביטוח בריאות'),
      bullet('הצגה ברורה ומשווה — אזרח ופקיד רואים את שתי האפשרויות זו לצד זו'),
      bullet('אפיון מלא מאושר — הלוגיקה העסקית מגובשת ומתועדת'),
      bullet('שלד UI קיים — מוכן לפיתוח מהיר'),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'423',   label:'אזרחים בשנה',    sub:'שצריכים לבחור' },
      { num:'100%',  label:'אפיון מוכן',     sub:'לוגיקה מגובשת' },
      { num:'3–5ימ', label:'לפרודקשן',      sub:'פיתוח ממוקד' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('פיתוח מלא ובדיקות עם פקידי שירות לפני פריסה.', { size:18 })]),
      spacer(10),
      bullet('אישור לפיתוח מלא'),
      bullet('בדיקות עם 5 פקידי שירות'),
      bullet('אישור משפטי-מקצועי על נוסחאות החישוב'),
      bullet('תיאום עם מחלקת IT לשילוב במערכות'),
      spacer(10),
      p([
        r('קריטי: ', { bold:true, size:18, color:'8B0000' }),
        r('כל יום ללא הכלי = אזרח שמקבל החלטה חיים ללא ייעוץ נכון.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar(),
  ]);
  await save(doc, OUT + '09-bchirat-gmala.docx');
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT 7: עוזר אישי חכם
// ═══════════════════════════════════════════════════════════════════════════════
async function buildOzerIshi() {
  const doc = buildDoc([
    topBar(),
    spacer(18),
    titleBlock(
      '🧠',
      'עוזר אישי חכם לאזרח',
      `ליווי פרואקטיבי של אזרחים ותיקים במיצוי זכויות — מבוסס ${eng('TailorPath')}`,
      `◉ MVP בזקנה — פיילוט עם 500 אזרחים, תקציב מאושר`
    ),
    spacer(22),

    sectionHeader('🎯  האתגר'),
    spacer(10),
    contentBox([
      p([r('אזרחים ותיקים לא ממצים זכויות בשל מורכבות בירוקרטית. אלמנות וזקנים בתחילת ירידה תפקודית נופלים בין הכיסאות.', { size:18 })]),
      spacer(14),
      bullet('מידע מפוזר בין גופים — ביטוח לאומי, משרדי ממשלה, עמותות'),
      bullet('אין ליווי אישי רציף לאורך שינויי חיים'),
      bullet('המערכת מגיבה — לא יוזמת'),
    ]),
    spacer(22),

    sectionHeader('💡  הפתרון'),
    spacer(10),
    contentBox([
      p([r(`עוזר אישי דיגיטלי מבוסס ${eng('AI')} — ליווי פרואקטיבי דרך ${eng('WhatsApp')}.`, { size:18 })]),
      spacer(14),
      bullet('זיהוי אוטומטי של זכויות לא ממומשות לפי פרופיל אישי'),
      bullet('ניווט מותאם בין גמלאות, שירותים והטבות'),
      bullet(`MVP בזקנה — אלמנות וזקנים. בהמשך: שיקום, אבטלה ועוד`),
      bullet(`פלטפורמת ${eng('TailorPath')} — תשתית מוכחת לליווי אישי`),
      spacer(14),
      p([r('999K₪ מאושרים, 80% מימון קרנות ביל"א — הפרויקט מגובה ומאושר.', { size:18, bold:true })]),
    ]),
    spacer(22),

    sectionHeader('📊  מדדים מרכזיים'),
    spacer(12),
    kpiRow([
      { num:'999K',      label:'₪ תקציב',         sub:'מאושר ומתוקצב' },
      { num:'500',       label:'אזרחי פיילוט',    sub:'זקנה ואלמנות' },
      { num:'WhatsApp',  label:'ערוץ ראשי',       sub:'נגיש לכולם' },
    ]),
    spacer(22),

    sectionHeader('🚀  השלב הבא'),
    spacer(10),
    ctaBox([
      p([r('כל הרכיבים מוכנים — פיילוט בהמתנה לקבלת החלטה על הפעלה.', { size:18 })]),
      spacer(10),
      bullet(`השלמת ${eng('MVP')} בזקנה — 500 אזרחים ותיקים`),
      bullet('הרחבה לתחומי שיקום ואבטלה'),
      bullet('חיבור לנתוני DWH לזיהוי פרואקטיבי של זכאויות'),
      spacer(10),
      p([
        r('החלטה נדרשת: ', { bold:true, size:18, color:'1B3A6B' }),
        r('אישור להפעלת הפיילוט — תקציב ואישור ארגוני מוכנים, מחכים לאות.', { size:18 }),
      ]),
    ]),
    spacer(20),
    footerBar(),
  ]);
  await save(doc, OUT + '10-ozer-ishi.docx');
}

// ─── Run all ──────────────────────────────────────────────────────────────────
(async () => {
  console.log('\nבונה 7 מסמכי DOCX נוספים...\n');
  await buildArnak();
  await buildMachshevon();
  await buildMitzui();
  await buildVaadot();
  await buildNursing();
  await buildBchirat();
  await buildOzerIshi();
  console.log('\n✅ 7 מסמכים נוצרו בהצלחה ב:', OUT);
})();
