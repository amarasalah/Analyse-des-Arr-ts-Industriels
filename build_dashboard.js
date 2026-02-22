const fs = require('fs');
const preloaded = fs.readFileSync('preloaded_data.js', 'utf8');
const dataStr = preloaded.replace('const PRELOADED=', '').replace(/;\s*$/, '');
let html = fs.readFileSync('dashboard_template.html', 'utf8');
// Replace the null placeholder + comment with real data
html = html.replace(/null;\s*\/\/\s*%%PRELOADED_DATA%%/, dataStr + ';');
fs.writeFileSync('dashboard.html', html, 'utf8');
console.log('Dashboard built! Size:', (html.length / 1024).toFixed(1), 'KB');
console.log('Has real data:', html.includes('Changement du moule'));
console.log('Has placeholder:', html.includes('%%'));
