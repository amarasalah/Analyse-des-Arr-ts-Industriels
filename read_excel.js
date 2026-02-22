const XLSX = require('xlsx');
const fs = require('fs');

const result = {};
const files = fs.readdirSync('.').filter(f => f.endsWith('.xlsx'));
console.log('Found Excel files:', files);

files.forEach(file => {
    const key = file.replace('.xlsx', '');
    const wb = XLSX.readFile(file);
    result[key] = {};
    wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        result[key][name] = data;
    });
});

fs.writeFileSync('excel_data.json', JSON.stringify(result, null, 2), 'utf8');
console.log('Done! Data written to excel_data.json');

// Print summary
for (const [file, sheets] of Object.entries(result)) {
    console.log(`\nFile: ${file}`);
    for (const [sheet, rows] of Object.entries(sheets)) {
        console.log(`  Sheet: ${sheet}, Rows: ${rows.length}`);
        if (rows.length > 0) console.log(`  Headers: ${JSON.stringify(rows[0])}`);
        if (rows.length > 1) console.log(`  First data row: ${JSON.stringify(rows[1])}`);
        if (rows.length > 2) console.log(`  Second data row: ${JSON.stringify(rows[2])}`);
    }
}
