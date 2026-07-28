#!/usr/bin/env node
/**
 * d587_to_dbc.mjs — D587 通信矩阵 → CANoe DBC 转换器
 *
 * 从 D587 平台化通讯协议 Excel 生成 ANSI (Windows-1252) 编码的 DBC 文件。
 * 无需外部依赖 (仅使用 Node.js 内置模块 + fast-xml-parser)。
 *
 * 用法: node d587_to_dbc.mjs <input.xlsx> [-o output.dbc]
 */

import { readFileSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'fs';
import { join } from 'path';
import { XMLParser } from 'fast-xml-parser';
import { execSync } from 'child_process';

// ─────────────────────────────────────────────────────────────
// 配置
// ─────────────────────────────────────────────────────────────

const SHEET_NAME = '5_Matrix';

// 列映射 (1-based)
const COL = {
    subnet: 'A',      // sub-net
    sender: 'B',      // Sender
    sigName: 'C',     // Signal Name
    sigNameCn: 'D',   // Name In Chinese
    id: 'E',          // ID (hex without 0x prefix)
    cycle: 'F',       // Cycle Time [ms]
    startBit: 'G',    // signal Start bit (bit)
    bitLen: 'H',      // signal Length (bit)
    factor: 'I',      // Factor
    offset: 'J',      // Offset
    pmin: 'K',        // Physical Range Min (dec)
    pmax: 'L',        // Physical Range Max (dec)
    sigType: 'M',     // Signal Type
    unit: 'N',        // Unit
    valDesc: 'O',     // Description (hex)
    initialVal: 'P',  // Initial Value (hex)
    note1: 'Q',       // 备注1
    invalidVal: 'R',  // Invalid Value (hex)
    routeMark: 'S',   // Route mark
    note: 'T',        // Note
    receiver: 'U',    // Receiver
};

// ─────────────────────────────────────────────────────────────
// XLSX解析
// ─────────────────────────────────────────────────────────────

function parseXlsx(filepath) {
    const tmpDir = join(process.cwd(), '_tmp_xlsx_dbc');
    if (!existsSync(tmpDir)) mkdirSync(tmpDir, { recursive: true });

    // Extract needed XML files
    execSync(
        `unzip -o "${filepath}" "xl/workbook.xml" "xl/_rels/workbook.xml.rels" "xl/sharedStrings.xml" "xl/worksheets/sheet5.xml" -d "${tmpDir}"`,
        { stdio: 'pipe' }
    );

    const parser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '@_',
        textNodeName: '#text',
        isArray: (name) => ['row', 'c'].includes(name),
    });

    // Parse shared strings
    const ssXml = readFileSync(join(tmpDir, 'xl/sharedStrings.xml'), 'utf-8');
    const ssp = parser.parse(ssXml);
    let sharedStrings = [];
    if (ssp.sst && ssp.sst.si) {
        const siList = Array.isArray(ssp.sst.si) ? ssp.sst.si : [ssp.sst.si];
        sharedStrings = siList.map(si => {
            if (si.t) return typeof si.t === 'object' ? (si.t['#text'] || '') : si.t;
            if (si.r) {
                const ts = Array.isArray(si.r)
                    ? si.r.map(r => r.t?.['#text'] || r.t || '').join('')
                    : (si.r.t?.['#text'] || si.r.t || '');
                return ts;
            }
            return '';
        });
    }

    // Parse sheet5
    const shXml = readFileSync(join(tmpDir, 'xl/worksheets/sheet5.xml'), 'utf-8');
    const sh = parser.parse(shXml);
    const rows = sh.worksheet.sheetData.row;

    // Build cell access map
    function getCell(rowNum, col) {
        const row = rows[rowNum - 1];
        if (!row || !row.c) return '';
        const cells = Array.isArray(row.c) ? row.c : [row.c];
        for (const cell of cells) {
            const ref = cell['@_r'];
            const cellCol = ref.replace(/[0-9]/g, '');
            if (cellCol === col) {
                return getCellValue(cell, sharedStrings);
            }
        }
        return '';
    }

    // Clean up temp
    rmSync(tmpDir, { recursive: true, force: true });

    return { getCell, totalRows: rows.length };
}

function getCellValue(cell, sharedStrings) {
    if (!cell) return '';
    const type = cell['@_t'] || '';
    const v = cell.v;
    if (v === undefined || v === null) return '';
    const raw = typeof v === 'object' ? (v['#text'] || '') : String(v);
    if (type === 's') {
        const idx = parseInt(raw);
        return idx < sharedStrings.length ? sharedStrings[idx] : '';
    }
    return raw;
}

// ─────────────────────────────────────────────────────────────
// 辅助函数
// ─────────────────────────────────────────────────────────────

function sanitizeNode(name) {
    name = safeStr(name);
    return name.replace(/[^A-Za-z0-9_]/g, '_')
               .replace(/_+/g, '_')
               .replace(/^_|_$/g, '');
}

function sanitizeSignal(name) {
    name = safeStr(name);
    return name.replace(/[^A-Za-z0-9_]/g, '_')
               .replace(/_+/g, '_')
               .replace(/^_|_$/g, '');
}

function cleanForAnsi(text) {
    let cleaned = '';
    for (const ch of text) {
        if (ch.charCodeAt(0) < 256) {
            cleaned += ch;
        } else {
            cleaned += ' ';
        }
    }
    return cleaned.replace(/  +/g, ' ').trim();
}

function safeStr(v) {
    return (v === undefined || v === null) ? '' : String(v);
}

function parseValDesc(text) {
    const valDefs = {};
    text = safeStr(text);
    if (!text || !text.trim()) return valDefs;
    for (const line of text.split('\n')) {
        const trimmed = line.trim();
        // Support both : and ： (full-width colon)
        const colonIdx = trimmed.indexOf(':');
        const fcolonIdx = trimmed.indexOf('：');
        const useIdx = (fcolonIdx >= 0 && (colonIdx < 0 || fcolonIdx < colonIdx)) ? fcolonIdx : colonIdx;
        if (useIdx < 0) continue;

        const valStr = trimmed.substring(0, useIdx).trim();
        let descStr = trimmed.substring(useIdx + 1).trim();
        // Truncate to 60 chars
        if (descStr.length > 60) descStr = descStr.substring(0, 60);
        // Handle range values like "0x1~0xC9:0%~100%"
        if (valStr.includes('~') || valStr.includes('～')) {
            // Range values can't be expressed in DBC VAL_ blocks, skip
            continue;
        }
        if (valStr.startsWith('0x') || valStr.startsWith('0X')) {
            try {
                valDefs[parseInt(valStr, 16)] = descStr;
            } catch { /* ignore */ }
        } else if (/^\d+$/.test(valStr)) {
            valDefs[parseInt(valStr, 10)] = descStr;
        }
    }
    return valDefs;
}

function resolveByteOrder(text) {
    const t = safeStr(text).toLowerCase();
    if (t.includes('intel')) return '@0+';
    return '@1+';  // Default Motorola
}

function parseCycleTime(raw) {
    raw = safeStr(raw);
    if (!raw) return 0;
    // Format: "500/20" or "500" or "100"
    const parts = raw.split('/');
    const val = parseFloat(parts[0].trim());
    return isNaN(val) ? 0 : Math.round(val);
}

function parseNumeric(raw, defaultVal) {
    raw = safeStr(raw);
    if (!raw || raw === '.' || raw.toLowerCase() === 'na' || raw.toLowerCase() === 'nan') return defaultVal;
    const val = parseFloat(raw);
    return isNaN(val) ? defaultVal : val;
}

function parseIntNumeric(raw, defaultVal) {
    raw = safeStr(raw);
    if (!raw || raw === '.' || raw.toLowerCase() === 'na' || raw.toLowerCase() === 'nan') return defaultVal;
    const val = parseInt(raw, 10);
    return isNaN(val) ? defaultVal : val;
}

function parseHexId(raw) {
    raw = safeStr(raw);
    if (!raw) return 0;
    const trimmed = raw.trim();
    // Use hex parsing (the ID column contains hex without 0x prefix)
    const val = parseInt(trimmed, 16);
    if (!isNaN(val)) return val;
    // Fallback to decimal
    const decVal = parseInt(trimmed, 10);
    return isNaN(decVal) ? 0 : decVal;
}

function extractRouteName(routeMark) {
    routeMark = safeStr(routeMark);
    if (!routeMark) return '';
    // Take first line if multiline
    const firstLine = routeMark.split('\n')[0].trim();
    // Route mark format: "B_SRS_050" or "I_HMDC/CDC_324"
    // The last part after the last underscore is the hex ID
    const lastUnderscore = firstLine.lastIndexOf('_');
    if (lastUnderscore < 0) return firstLine;
    const prefix = firstLine.substring(0, lastUnderscore);
    return prefix || firstLine;
}

function extractRouteId(routeMark) {
    routeMark = safeStr(routeMark);
    if (!routeMark) return null;
    const firstLine = routeMark.split('\n')[0].trim();
    const lastUnderscore = firstLine.lastIndexOf('_');
    if (lastUnderscore < 0) return null;
    const suffix = firstLine.substring(lastUnderscore + 1);
    // The suffix should be a hex number
    if (/^[0-9A-Fa-f]+$/.test(suffix)) {
        return parseInt(suffix, 16);
    }
    return null;
}

function extractSenders(raw) {
    raw = safeStr(raw);
    if (!raw) return [];
    // Some senders are like "RSML/RSM/LSCM" - split by / or ,
    const parts = raw.split(/[\/,]/).map(s => sanitizeNode(s.trim())).filter(Boolean);
    // Deduplicate
    return [...new Set(parts)];
}

function extractReceivers(raw) {
    raw = safeStr(raw);
    if (!raw) return [];
    const parts = raw.split(',').map(s => sanitizeNode(s.trim())).filter(Boolean);
    return [...new Set(parts)];
}

// ─────────────────────────────────────────────────────────────
// 主转换逻辑
// ─────────────────────────────────────────────────────────────

function convert(inputPath) {
    const { getCell, totalRows } = parseXlsx(inputPath);
    console.log(`已解析 ${totalRows} 行`);

    // Group signals by (Sender, CAN ID)
    const groups = new Map();

    for (let r = 2; r <= totalRows; r++) {
        const sigName = getCell(r, COL.sigName);
        const idRaw = getCell(r, COL.id);
        if (!sigName || !idRaw) continue;

        const canId = parseHexId(idRaw);
        if (canId === 0) continue;

        const senderRaw = getCell(r, COL.sender);
        if (!senderRaw) continue;

        const senders = extractSenders(senderRaw);
        const primarySender = senders[0] || 'Vector__XXX';

        // Use primary sender for grouping
        const key = `${primarySender}|${canId}`;
        if (!groups.has(key)) {
            groups.set(key, []);
        }
        groups.get(key).push({ row: r, canId, senderRaw, primarySender, senders, sigName });
    }

    console.log(`发现 ${groups.size} 个报文组`);

    // Process each group into a frame
    const frames = [];
    const usedMsgNames = new Set();

    for (const [key, signalRows] of groups) {
        const [sender, idStr] = key.split('|');
        const canId = parseInt(idStr);
        if (canId === 0) continue;

        // Collect sender and receiver info
        const allSenders = new Set();
        const allReceivers = new Set();
        let cycleTime = 0;
        let routeMark = '';

        for (const sr of signalRows) {
            const r = sr.row;
            // Senders
            const senders = extractSenders(getCell(r, COL.sender));
            senders.forEach(s => allSenders.add(s));
            // Receivers
            const rcvs = extractReceivers(getCell(r, COL.receiver));
            rcvs.forEach(s => allReceivers.add(s));
            // Cycle time (use the first non-zero value)
            const ct = parseCycleTime(getCell(r, COL.cycle));
            if (ct > 0 && cycleTime === 0) cycleTime = ct;
            // Route mark
            const rm = getCell(r, COL.routeMark);
            if (rm && !routeMark) routeMark = rm;
        }

        // Determine message name from route mark prefix + actual CAN ID
        let msgName = '';
        const hexId = canId.toString(16).toUpperCase().padStart(3, '0');
        if (routeMark) {
            const firstLine = routeMark.split('\n')[0].trim();
            const prefix = extractRouteName(firstLine);  // Everything before last underscore
            if (prefix) {
                msgName = `${sanitizeSignal(prefix)}_${hexId}`;
            }
        }
        if (!msgName) {
            // No route mark or empty prefix: generate from sender and ID
            msgName = `${sanitizeNode([...allSenders][0] || 'ECU')}_${hexId}`;
        }
        if (!msgName) {
            // Generate from sender and ID
            msgName = `${sender}_${canId.toString(16).toUpperCase().padStart(3, '0')}`;
        }

        // Primary sender
        const msgSender = sanitizeNode([...allSenders][0] || sender);

        // Calculate DLC from signals
        let maxBit = -1;
        const signals = [];

        for (const sr of signalRows) {
            const r = sr.row;
            const sigName = getCell(r, COL.sigName);
            const startBitRaw = getCell(r, COL.startBit);
            const bitLenRaw = getCell(r, COL.bitLen);
            const factorRaw = getCell(r, COL.factor);
            const offsetRaw = getCell(r, COL.offset);
            const pminRaw = getCell(r, COL.pmin);
            const pmaxRaw = getCell(r, COL.pmax);
            const sigTypeRaw = getCell(r, COL.sigType);
            const unitRaw = getCell(r, COL.unit);
            const valDescRaw = getCell(r, COL.valDesc);
            const descCn = getCell(r, COL.sigNameCn);

            const safeName = sanitizeSignal(sigName);
            const startBit = parseIntNumeric(startBitRaw, 0);
            const bitLen = parseIntNumeric(bitLenRaw, 1);
            const factor = parseNumeric(factorRaw, 1);
            const offset = parseNumeric(offsetRaw, 0);
            const pmin = parseNumeric(pminRaw, 0);
            const pmax = parseNumeric(pmaxRaw, pmin + (1 << bitLen) - 1);
            const unit = unitRaw;

            // Determine byte order and signedness
            // No byte order column in this matrix - default to Motorola
            // Signal Type is "unsigned" or "Unsigned"
            // Signal Type: check for signed (but not unsigned)
            const sigTypeLower = safeStr(sigTypeRaw).toLowerCase();
            const isSigned = sigTypeLower === 'signed' || sigTypeLower === 'intel' || sigTypeLower === 'signed_motorola';
            const bo = isSigned ? '@1-' : '@1+';  // Motorola

            const maxPos = startBit + bitLen - 1;
            if (maxPos > maxBit) maxBit = maxPos;

            // Signal description: prefer English signal name as comment
            // Chinese description is used when it provides additional info (has Latin-1 content)
            const sigDesc = (descCn && cleanForAnsi(descCn)) ? descCn : sigName;

            // VAL_ definitions
            const valDefs = parseValDesc(valDescRaw);

            // Signal receivers (per signal level)
            const sigRcvRaw = getCell(r, COL.receiver);
            const sigRcv = extractReceivers(sigRcvRaw);

            signals.push({
                name: safeName,
                startBit,
                bitLen,
                factor,
                offset,
                pmin,
                pmax,
                unit,
                description: sigDesc,
                byteOrder: bo,
                valDefs,
                receivers: sigRcv,
                isSigned,
                origName: sigName,
            });
        }

        const dlc = maxBit >= 0 ? Math.floor(maxBit / 8) + 1 : 8;
        // Clamp DLC to 0-8 for CAN (or 0-64 for CAN-FD, but 8 is safe)
        const clampedDlc = Math.min(Math.max(dlc, 1), 64);

        frames.push({
            id: canId,
            name: msgName,
            dlc: clampedDlc,
            cycle: cycleTime,
            sender: msgSender,
            allSenders: [...allSenders],
            allReceivers: [...allReceivers],
            signals,
            routeMark,
        });
    }

    // Sort frames by CAN ID
    frames.sort((a, b) => a.id - b.id);

    return frames;
}

// ─────────────────────────────────────────────────────────────
// DBC 生成
// ─────────────────────────────────────────────────────────────

function genDbc(frames) {
    const lines = [];
    lines.push('VERSION ""');
    lines.push('');

    // NS_
    lines.push('NS_ :');
    const nsItems = [
        'NS_DESC_', 'CM_', 'BA_DEF_', 'BA_DEF_DEF_',
        'BA_DEF_DEF_REL_', 'BA_DEF_SGTYPE_', 'BA_DEF_REL_',
        'BA_DEF_DEF_', 'BU_DEF_REL_', 'BA_', 'BA_DEF_',
        'VAL_', 'CAT_DEF_', 'CAT_', 'FILTER',
        'BA_DEF_DEF_', 'EV_DATA_', 'ENVVAR_DATA_', 'SGTYPE_',
        'SGTYPE_VAL_', 'BA_DEF_SGTYPE_', 'BA_SGTYPE_', 'SIG_VALTYPE_',
        'SIGTYPE_VALTYPE_', 'BO_TX_BU_', 'BA_DEF_REL_',
        'BA_DEF_DEF_REL_', 'BA_REL_', 'BA_DEF_DEF_', 'BU_BO_REL_',
        'BA_DEF_', 'BU_SGP_REL_', 'BU_SGT_REL_', 'SG_MUL_VAL_',
    ];
    for (const ns of nsItems) {
        lines.push(`  ${ns}`);
    }
    lines.push('');

    // BS_
    lines.push('BS_:');
    lines.push('');

    // BU_ - collect all nodes
    const allNodes = new Set();
    for (const f of frames) {
        allNodes.add(sanitizeNode(f.sender));
        for (const s of f.allSenders) allNodes.add(sanitizeNode(s));
        for (const r of f.allReceivers) allNodes.add(sanitizeNode(r));
        for (const sig of f.signals) {
            for (const r of sig.receivers) allNodes.add(sanitizeNode(r));
        }
    }
    allNodes.delete(''); // Remove empty string if any
    lines.push(`BU_: ${[...allNodes].sort().join(' ')}`);
    lines.push('');

    // BO_ + SG_
    for (const f of frames) {
        const sender = sanitizeNode(f.sender);
        const sigLines = [];
        for (const sig of f.signals) {
            // Determine receivers for this signal
            let rcvStr = '';
            if (sig.receivers && sig.receivers.length > 0) {
                rcvStr = sig.receivers.map(s => sanitizeNode(s)).join(',');
            } else if (f.allReceivers.length > 0) {
                rcvStr = f.allReceivers.map(s => sanitizeNode(s)).join(',');
            } else {
                rcvStr = 'Vector__XXX';
            }

            const minV = sig.pmin;
            const maxV = sig.pmax > sig.pmin ? sig.pmax : sig.pmin + (1 << sig.bitLen) - 1;

            sigLines.push(
                `  SG_ ${sig.name} : ${sig.startBit}|${sig.bitLen}${sig.byteOrder} (${sig.factor},${sig.offset}) [${minV}|${maxV}] "${sig.unit}"  ${rcvStr}`
            );
        }
        lines.push(`BO_ ${f.id} ${f.name}: ${f.dlc} ${sender}`);
        lines.push(...sigLines);
        lines.push('');
    }

    // CM_ SG_
    for (const f of frames) {
        for (const sig of f.signals) {
            if (sig.description) {
                const desc = sig.description.replace(/"/g, "'");
                const descClean = cleanForAnsi(desc);
                if (descClean) {
                    lines.push(`CM_ SG_ ${f.id} ${sig.name} "${descClean}";`);
                }
            }
        }
    }
    if (lines[lines.length - 1] !== '') lines.push('');

    // CM_ BO_
    for (const f of frames) {
        const cmt = f.routeMark ? f.routeMark.split('\n')[0].trim() : f.name;
        const cmtClean = cleanForAnsi(cmt.replace(/"/g, "'")).replace(/\//g, '_');
        if (cmtClean) {
            lines.push(`CM_ BO_ ${f.id} "${cmtClean}";`);
        }
    }
    lines.push('');

    // BA_DEF_
    lines.push('BA_DEF_ BO_ "GenMsgCycleTime" INT 0 65535;');
    lines.push('BA_DEF_ BO_ "GenMsgSendType" STRING;');
    lines.push('');
    lines.push('BA_DEF_DEF_ "GenMsgCycleTime" 0;');
    lines.push('BA_DEF_DEF_ "GenMsgSendType" "Cyclic";');
    lines.push('');

    // BA_ attributes
    for (const f of frames) {
        lines.push(`BA_ "GenMsgCycleTime" BO_ ${f.id} ${f.cycle};`);

        // Determine send type based on cycle format (if cycle has "/", it's event-triggered)
        // We don't have the exact send type per message, so default to "Cyclic"
        lines.push(`BA_ "GenMsgSendType" BO_ ${f.id} "Cyclic";`);
    }
    lines.push('');

    // VAL_
    for (const f of frames) {
        for (const sig of f.signals) {
            if (Object.keys(sig.valDefs).length > 0) {
                for (const [k, v] of Object.entries(sig.valDefs)) {
                    const vDesc = v.replace(/"/g, "'");
                    lines.push(`VAL_ ${f.id} ${sig.name} ${k} "${vDesc}" ;`);
                }
            } else if (sig.bitLen === 1 && !sig.name.toLowerCase().includes('chks') && !sig.name.toLowerCase().includes('chk')) {
                // 1-bit signals get Active/Inactive enum (skip checksums)
                lines.push(`VAL_ ${f.id} ${sig.name} 1 "Active" 0 "Inactive" ;`);
            }
        }
    }
    lines.push('');

    return lines.join('\n');
}

function saveDbc(dbcText, outputPath) {
    // Clean for ANSI encoding
    const clean = cleanForAnsi(dbcText);
    // Node.js doesn't support windows-1252 natively.
    // latin1 is close enough for DBC files (only differs in 0x80-0x9F range).
    // For maximum compatibility, manually encode each character.
    const encoder = new TextEncoder('utf-8');
    // Encode each Latin-1 character to a byte
    const bytes = new Uint8Array(clean.length);
    for (let i = 0; i < clean.length; i++) {
        const code = clean.charCodeAt(i);
        if (code < 256) {
            bytes[i] = code;
        } else {
            bytes[i] = 0x20; // Replace non-Latin1 with space
        }
    }
    writeFileSync(outputPath, Buffer.from(bytes));
}

// ─────────────────────────────────────────────────────────────
// 摘要输出
// ─────────────────────────────────────────────────────────────

function printSummary(frames, outputPath) {
    const totalSigs = frames.reduce((sum, f) => sum + f.signals.length, 0);
    console.log(`\nDBC 文件: ${outputPath}`);
    console.log(`帧数: ${frames.length}`);
    console.log(`信号数: ${totalSigs}`);
    console.log(`节点数: ${new Set(frames.flatMap(f => [f.sender, ...f.allSenders, ...f.allReceivers])).size}`);
    console.log();
    console.log(`${'ID'.padStart(8)}  ${'Frame Name'.padEnd(45)} ${'DLC'.padStart(4)} ${'Cycle'.padStart(6)}  ${'Sender'.padEnd(12)}  Signals`);
    console.log('-'.repeat(95));
    for (const f of frames) {
        const idStr = `0x${f.id.toString(16).toUpperCase().padStart(3, '0')}`;
        console.log(`${idStr.padStart(8)}  ${f.name.padEnd(45)} ${String(f.dlc).padStart(4)} ${String(f.cycle).padStart(6)}  ${f.sender.padEnd(12)}  ${f.signals.length}`);
    }
    console.log();
}

// ─────────────────────────────────────────────────────────────
// 入口
// ─────────────────────────────────────────────────────────────

function main() {
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('用法: node d587_to_dbc.mjs <input.xlsx> [-o output.dbc]');
        process.exit(1);
    }

    const inputPath = args[0];
    let outputPath;
    const oIdx = args.indexOf('-o');
    if (oIdx >= 0 && oIdx + 1 < args.length) {
        outputPath = args[oIdx + 1];
    } else {
        const base = inputPath.replace(/\.xlsx$/i, '');
        outputPath = base + '.dbc';
    }

    console.log(`输入: ${inputPath}`);
    console.log(`输出: ${outputPath}\n`);

    const frames = convert(inputPath);

    if (frames.length === 0) {
        console.log('错误: 没有解析到任何帧');
        process.exit(1);
    }

    const dbcText = genDbc(frames);
    saveDbc(dbcText, outputPath);
    printSummary(frames, outputPath);
    console.log('生成完成!');
}

main();
