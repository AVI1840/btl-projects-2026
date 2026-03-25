/**
 * Script to update all one-pager HTML files:
 * 1. Change button text from "השאירו משוב" to "משוב לשיפור"
 * 2. Add Google Sheets integration to fbSubmit
 * 3. Change "שמור משוב" to "שלח משוב"
 * 4. Update hint text
 */
const fs = require('fs');
const path = require('path');

const htmlDir = path.join(__dirname, '..', 'html');
const SHEET_URL = 'https://script.google.com/macros/s/AKfycbwD8CMFoP5XoOwRLwK_OxMMOFKF8fS2CRpbJkNdOHjbnJIepkOLzlGrg3GQNGRqbwB6bA/exec';

// Files to update (skip index.html, feedback-dashboard.html, btl_page.html, arnak already done)
const files = fs.readdirSync(htmlDir)
  .filter(f => f.endsWith('.html'))
  .filter(f => !['index.html', 'feedback-dashboard.html', 'btl_page.html', 'arnak-zchuyot.html'].includes(f));

let updated = 0;

for (const file of files) {
  const filePath = path.join(htmlDir, file);
  let html = fs.readFileSync(filePath, 'utf8');

  // Skip files without feedback system
  if (!html.includes('fb-btn') || !html.includes('fbSubmit')) {
    console.log(`⏭  ${file} — no feedback system, skipping`);
    continue;
  }

  // Extract FK (localStorage key) and APP_NAME from fb-title
  const fkMatch = html.match(/const FK\s*=\s*'([^']+)'/);
  const titleMatch = html.match(/fb-title[^>]*>💬\s*משוב[^—]*—\s*([^<]+)</);
  if (!fkMatch) {
    console.log(`⚠  ${file} — can't find FK, skipping`);
    continue;
  }
  const FK = fkMatch[1];
  const appName = titleMatch ? titleMatch[1].trim() : file.replace('.html', '');

  // 1. Button text: "השאירו משוב" → "משוב לשיפור"
  html = html.replace(/💬 השאירו משוב/g, '💬 משוב לשיפור');

  // 2. Panel title: ensure "משוב לשיפור"
  html = html.replace(/💬 משוב —/g, '💬 משוב לשיפור —');

  // 3. Submit button: "שמור משוב" → "שלח משוב"
  html = html.replace(/✅ שמור משוב/g, '✅ שלח משוב');

  // 4. Hint text update
  html = html.replace(
    /הפידבק נשמר בדפדפן זה ✓ לייצוא לחץ על "היסטוריה"/g,
    'הפידבק נשמר מקומית ונשלח לגוגל שיטס ✓'
  );

  // 5. Add SHEET_URL and APP_NAME constants + sendToSheet function after FK line
  if (!html.includes('SHEET_URL')) {
    html = html.replace(
      `const FK = '${FK}';`,
      `const FK = '${FK}';\nconst APP_NAME = '${appName}';\nconst SHEET_URL = '${SHEET_URL}';\nasync function sendToSheet(text, category) { try { await fetch(SHEET_URL, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ app: APP_NAME, name: 'one-pager', category, severity: '—', text, page: window.location.pathname }) }); } catch {} }`
    );
  }

  // 6. Make fbSubmit async and add sendToSheet call
  // Replace "function fbSubmit()" with "async function fbSubmit()"
  html = html.replace(/function fbSubmit\(\)/g, 'async function fbSubmit()');

  // Add sendToSheet call before localStorage save — inject after the validation check
  // We look for the pattern after the "if (!pro.r && ..." check, before "const entry ="
  if (!html.includes('sendToSheet(')) {
    html = html.replace(
      /if \(!pro\.r && !proc\.r && !des\.r && !pro\.t && !proc\.t && !des\.t\) \{[^}]+\}\n/,
      (match) => {
        return match +
          `  const parts = [];\n` +
          `  if (pro.t || pro.r) parts.push('מקצועי (' + pro.r + '★): ' + pro.t);\n` +
          `  if (proc.t || proc.r) parts.push('תהליכי (' + proc.r + '★): ' + proc.t);\n` +
          `  if (des.t || des.r) parts.push('עיצובי (' + des.r + '★): ' + des.t);\n` +
          `  await sendToSheet(parts.join(' | '), pro.t ? '🏛 מקצועי' : proc.t ? '⚙ תהליכי' : '🎨 עיצובי');\n`;
      }
    );
  }

  // 7. Change alert to status message if possible (simpler: just update alert text)
  html = html.replace(
    /alert\('✅ תודה! הפידבק נשמר\.'\)/g,
    "alert('✅ תודה! המשוב נשלח בהצלחה.')"
  );

  fs.writeFileSync(filePath, html, 'utf8');
  console.log(`✅ ${file} — updated (FK: ${FK}, app: ${appName})`);
  updated++;
}

console.log(`\nDone! Updated ${updated} files.`);
