import { readFileSync, existsSync } from 'fs';
import { join } from 'path';
import { XMLParser } from 'fast-xml-parser';

const tmpDir = '_tmp_ldf';

const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    isArray: (name) => ['row', 'c'].includes(name),
});

// Parse shared strings
const ss = parser.parse(readFileSync(join(tmpDir, 'xl/sharedStrings.xml'), 'utf-8'));
let ssList = [];
if (ss.sst && ss.sst.si) {
    const arr = Array.isArray(ss.sst.si) ? ss.sst.si : [ss.sst.si];
    ssList = arr.map(si => {
        if (si.t) return typeof si.t === 'object' ? (si.t['#text'] || '') : si.t;
        if (si.r) {
            const ts = Array.isArray(si.r) ? si.r.map(r => r.t?.['#text'] || r.t || '').join('') : (si.r.t?.['#text'] || si.r.t || '');
            return ts;
        }
        return '';
    });
}

function cv(cell) {
    if (!cell) return '';
    const t = cell['@_t'] || '';
    const v = cell.v;
    if (v === undefined || v === null) return '';
    const raw = typeof v === 'object' ? (v['#text'] || '') : String(v);
    if (t === 's') {
        const idx = parseInt(raw);
        return idx < ssList.length ? ssList[idx] : '';
    }
    return raw;
}

function cl(ref) { return ref.replace(/[0-9]/g, ''); }

function dumpSheet(sheetIndex, label, maxRows) {
    const filepath = join(tmpDir, `xl/worksheets/sheet${sheetIndex}.xml`);
    if (!existsSync(filepath)) { console.log(`\n=== ${label} === (not found)`); return; }

    const sh = parser.parse(readFileSync(filepath, 'utf-8'));
    const data = sh.worksheet.sheetData;
    if (!data || !data.row) { console.log(`\n=== ${label} === (empty)`); return; }
    const rows = Array.isArray(data.row) ? data.row : [data.row];

    console.log(`\n=== ${label} (${rows.length} rows) ===`);
    for (let ri = 0; ri < Math.min(maxRows || 10, rows.length); ri++) {
        const row = rows[ri];
        const rn = row['@_r'];
        const cells = row.c;
        const cArr = Array.isArray(cells) ? cells : [cells];
        const vals = [];
        for (const c of cArr) {
            const v = cv(c);
            if (v !== undefined && v !== null && v !== '') {
                vals.push(cl(c['@_r']) + '="' + String(v).replace(/\n/g, '\\n').substring(0, 80) + '"');
            }
        }
        if (vals.length > 0) console.log('  R' + rn + ': ' + vals.join(', '));
    }
}

// Sheet mapping from workbook
const sheetNames = ['0_CoverPage','修改记录','调度表','节点唤醒与休眠','特别说明',
                    'DR_LIN1','MMA_LIN_1','MMAV1_LIN1','MMAV2_LIN1',
                    'RRL_LIN1','MRL_LIN_1','MRLV1_LIN_1','MRLV2_LIN_1'];

for (let i = 0; i < sheetNames.length; i++) {
    dumpSheet(i + 1, sheetNames[i], 20);
}

console.log('\nDone!');
