const XLSX = require('xlsx');
const fs = require('fs');

const files = fs.readdirSync('.').filter(f => f.endsWith('.xlsx'));
const allCauses = [];

function parseTime(str) {
    if (!str || typeof str !== 'string') return 0;
    let mins = 0;
    const hMatch = str.match(/(\d+)\s*h/i);
    const mMatch = str.match(/(\d+)\s*min/i);
    if (hMatch) mins += parseInt(hMatch[1]) * 60;
    if (mMatch) mins += parseInt(mMatch[1]);
    if (!hMatch && !mMatch && str.match(/(\d+)\s*heures/i)) {
        mins = parseInt(str.match(/(\d+)\s*heures/i)[1]) * 60;
    }
    return mins;
}

function categorize(cause) {
    const c = cause.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (c.includes('moule') && (c.includes('changement') || c.includes('montage') || c.includes('demontage'))) return 'Changement/Montage Moule';
    if (c.includes('moule') && (c.includes('soudure') || c.includes('reglage') || c.includes('serrage') || c.includes('nettoyage') || c.includes('dressage'))) return 'Réglage/Maintenance Moule';
    if (c.includes('malaxeur') || c.includes('nettoyage') && c.includes('lavage')) return 'Nettoyage Malaxeur';
    if (c.includes('nettoyage')) return 'Nettoyage Général';
    if (c.includes('verin') || c.includes('dame')) return 'Vérin/Dame';
    if (c.includes('tiroir')) return 'Tiroir';
    if (c.includes('compresseur') || c.includes('air')) return 'Problème Air/Compresseur';
    if (c.includes('chaine') || c.includes('extracteur') || c.includes('ascenseur')) return 'Chaîne/Extracteur/Ascenseur';
    if (c.includes('soudure')) return 'Soudure';
    if (c.includes('capteur') || c.includes('logiciel') || c.includes('electrique') || c.includes('armoire') || c.includes('variateur')) return 'Électrique/Capteur/Logiciel';
    if (c.includes('vibreur') || c.includes('courroie') || c.includes('cardan') || c.includes('presse')) return 'Vibreur/Presse';
    if (c.includes('bascule') || c.includes('tremie') || c.includes('skip') || c.includes('hydraulique')) return 'Bascule/Trémie/Hydraulique';
    if (c.includes('coincement') || c.includes('planche') || c.includes('blocage')) return 'Coincement/Blocage Planche';
    if (c.includes('flexible')) return 'Changement Flexible';
    if (c.includes('panne') || c.includes('reparation') || c.includes('probleme')) return 'Pannes Diverses';
    return 'Autres';
}

files.forEach(file => {
    const wb = XLSX.readFile(file);
    const month = file.replace('.xlsx', '');
    wb.SheetNames.forEach(sheetName => {
        const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
        let currentDate = null, currentTime = null;
        data.forEach(row => {
            if (row[0] && typeof row[0] === 'number' && row[0] > 40000) {
                const d = new Date((row[0] - 25569) * 86400 * 1000);
                currentDate = d.toISOString().split('T')[0];
                currentTime = row[1] ? String(row[1]).trim() : null;
            }
            const cause = row[3] ? String(row[3]).trim() : null;
            if (cause && !cause.startsWith('Causes') && cause.length > 3 && !cause.startsWith('Démarrage') && !cause.toLowerCase().includes('pause') && !cause.toLowerCase().includes('retard')) {
                allCauses.push({
                    file: month, sheet: sheetName, date: currentDate,
                    time: currentTime, timeMinutes: parseTime(currentTime),
                    cause: cause, category: categorize(cause)
                });
            }
        });
    });
});

// Build stats
const catCount = {}, catByMonth = {};
allCauses.forEach(c => {
    catCount[c.category] = (catCount[c.category] || 0) + 1;
    if (!catByMonth[c.file]) catByMonth[c.file] = {};
    catByMonth[c.file][c.category] = (catByMonth[c.file][c.category] || 0) + 1;
});

const total = allCauses.length;
const stats = Object.entries(catCount).map(([cat, count]) => ({
    category: cat, count, percentage: ((count / total) * 100).toFixed(1)
})).sort((a, b) => b.count - a.count);

const output = { totalEvents: total, categories: stats, byMonth: catByMonth, rawCauses: allCauses, files: files };
fs.writeFileSync('analysis.json', JSON.stringify(output, null, 2), 'utf8');
console.log(JSON.stringify(output, null, 2));
