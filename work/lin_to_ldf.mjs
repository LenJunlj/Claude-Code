#!/usr/bin/env node
/**
 * lin_to_ldf.mjs — LIN 通信矩阵 Excel → Vector LDF 转换器
 *
 * 从 D587 BDC_LIN 协议 Excel 生成 Vector LDF Explorer Pro 兼容的 .ldf 文件。
 * 使用 fast-xml-parser 直接解析 XLSX，无需 openpyxl。
 *
 * 用法: node lin_to_ldf.mjs <input.xlsx> [-o output.ldf]
 */

import { readFileSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'fs';
import { join } from 'path';
import { XMLParser } from 'fast-xml-parser';
import { execSync } from 'child_process';

// ─────────────────────────────────────────────────────────────
// 主函数
// ─────────────────────────────────────────────────────────────

function convert(inputPath) {
    const tmpDir = join(process.cwd(), '_tmp_ldf_gen');
    if (!existsSync(tmpDir)) mkdirSync(tmpDir, { recursive: true });

    execSync(
        `unzip -o "${inputPath}" "xl/workbook.xml" "xl/_rels/workbook.xml.rels" "xl/sharedStrings.xml" "xl/worksheets/sheet*.xml" -d "${tmpDir}"`,
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
                const ts = Array.isArray(si.r) ? si.r.map(r => r.t?.['#text'] || r.t || '').join('') : (si.r.t?.['#text'] || si.r.t || '');
                return ts;
            }
            return '';
        });
    }

    // Map sheet names to files
    const wbXml = readFileSync(join(tmpDir, 'xl/workbook.xml'), 'utf-8');
    const wb = parser.parse(wbXml);
    const sheetEntries = wb.workbook.sheets.sheet;
    const sheetNames = sheetEntries.map(s => s['@_name']);

    const relsXml = readFileSync(join(tmpDir, 'xl/_rels/workbook.xml.rels'), 'utf-8');
    const rels = parser.parse(relsXml);
    const relMap = {};
    for (const rel of (Array.isArray(rels.Relationships.Relationship) ? rels.Relationships.Relationship : [rels.Relationships.Relationship])) {
        relMap[rel['@_Id']] = rel['@_Target'];
    }

    // Build sheet index
    let sheetIndex = 1;
    const sheetFileMap = {};
    for (const sheet of sheetEntries) {
        const rid = sheet['@_r:id'];
        const target = relMap[rid];
        if (target) {
            const filePath = join(tmpDir, 'xl', target);
            sheetFileMap[sheet['@_name']] = filePath;
        }
    }

    // Helper to read a sheet
    function getSheet(name) {
        const fp = sheetFileMap[name];
        if (!fp || !existsSync(fp)) return null;
        const sh = parser.parse(readFileSync(fp, 'utf-8'));
        const data = sh.worksheet.sheetData;
        if (!data || !data.row) return { rows: [], getCell: () => '' };
        const rows = Array.isArray(data.row) ? data.row : [data.row];

        return {
            rows,
            getCell(rowNum, col) {
                // Find row by @_r attribute (not array index, since blank rows are omitted)
                for (const row of rows) {
                    if (parseInt(row['@_r']) === rowNum) {
                        if (!row.c) return '';
                        const cells = Array.isArray(row.c) ? row.c : [row.c];
                        for (const cell of cells) {
                            const colLetter = cell['@_r'].replace(/[0-9]/g, '');
                            if (colLetter === col) {
                                return getCellValue(cell, sharedStrings);
                            }
                        }
                        return '';
                    }
                }
                return '';
            }
        };
    }

    function getCellValue(cell, ss) {
        if (!cell) return '';
        const t = cell['@_t'] || '';
        const v = cell.v;
        if (v === undefined || v === null) return '';
        const raw = typeof v === 'object' ? (v['#text'] || '') : String(v);
        if (t === 's') {
            const idx = parseInt(raw);
            return idx < ss.length ? ss[idx] : '';
        }
        return raw;
    }

    function safeStr(v) { return (v === undefined || v === null) ? '' : String(v); }

    // ── Parse 调度表 (Schedule) ──
    const schedSheet = getSheet('调度表');
    let baudRate = '19.2';
    let scheduleCycle = 80;
    let timeBase = 10;

    if (schedSheet) {
        // Row 3: 调度周期
        const cycleStr = schedSheet.getCell(3, 'B');
        if (cycleStr) {
            const m = cycleStr.match(/(\d+)/);
            if (m) scheduleCycle = parseInt(m[1]);
        }
        // Row 4: 时间域量
        const tbStr = schedSheet.getCell(4, 'B');
        if (tbStr) {
            const m = tbStr.match(/(\d+)/);
            if (m) timeBase = parseInt(m[1]);
        }
        // Row 10: speed
        const speedStr = schedSheet.getCell(10, 'A');
        if (speedStr) {
            const m = speedStr.match(/([\d.]+)/);
            if (m) baudRate = m[1];
        }
    }

    // Or try 节点唤醒与休眠 sheet for speed
    const wakeSheet = getSheet('节点唤醒与休眠');
    if (wakeSheet) {
        const speedVal = wakeSheet.getCell(11, 'B');
        if (speedVal) {
            const m = speedVal.match(/(\d+)/);
            if (m) {
                const bps = parseInt(m[1]);
                baudRate = String(bps / 1000);
            }
        }
    }

    console.log(`波特率: ${baudRate} kbps, 调度周期: ${scheduleCycle}ms, 时间域量: ${timeBase}ms`);

    // ── Data sheets mapping (LIN bus definitions) ──
    const linSheets = ['DR_LIN1', 'MMA_LIN_1', 'MMAV1_LIN1', 'MMAV2_LIN1',
                       'RRL_LIN1', 'MRL_LIN_1', 'MRLV1_LIN_1', 'MRLV2_LIN_1'];

    // ── Parse each data sheet ──
    const frames = [];

    // Extract master/slave info from schedule
    // Row 6 in schedule: B=主节点, D=主节点, F=从节点, H=从节点, J=从节点, L=从节点, N=从节点, P=从节点
    const scheduleNodes = {};
    const schedCols = { B: 0, D: 1, F: 2, H: 3, J: 4, L: 5, N: 6, P: 7 };
    if (schedSheet) {
        for (const [col, idx] of Object.entries(schedCols)) {
            const nodeType = schedSheet.getCell(6, col);
            if (nodeType) scheduleNodes[idx] = nodeType;
        }
    }

    // Global signal name set for cross-frame uniqueness (name → Set of publishers)
    const globalSignalNames = new Map();

    for (let si = 0; si < linSheets.length; si++) {
        const sheetName = linSheets[si];
        const sheet = getSheet(sheetName);
        if (!sheet || sheet.rows.length < 14) {
            console.log(`  跳过: ${sheetName} (空或无数据)`);
            continue;
        }

        // Parse header rows
        const frameName = safeStr(sheet.getCell(2, 'C')).trim();
        const idStr = safeStr(sheet.getCell(3, 'C')).trim();
        const senderRaw = safeStr(sheet.getCell(4, 'C')).trim();
        const receiverRaw = safeStr(sheet.getCell(5, 'C')).trim();
        const lengthStr = safeStr(sheet.getCell(7, 'C')).trim();

        if (!frameName || !idStr) continue;

        const frameId = parseInt(idStr, 10);
        const msgLength = parseInt(lengthStr, 10) || 8;

        // Extract node names from Chinese descriptions
        // "左侧智能座椅终端（LSCM）" → LSCM
        // "主驾腰托按摩（LRM_D）" → LRM_D
        // "中排腰托按摩_左（LRM_ML）" → LRM_ML
        function extractNode(name) {
            const m = name.match(/[（\(]([^）\)]+)[）\)]/);
            return m ? m[1].trim() : name.replace(/[^A-Za-z0-9_]/g, '_');
        }

        const pubName = extractNode(senderRaw);
        const subName = extractNode(receiverRaw);

        console.log(`  ${frameName}: ID=${frameId}, 发布者=${pubName}, 订阅者=${subName}`);

        // Parse signal data rows (from row 14 onwards)
        const signals = [];
        const usedOriginalNamesInFrame = new Set();  // Track original names for frame-level dedup
        for (let r = 14; r <= sheet.rows.length; r++) {
            let sigName = safeStr(sheet.getCell(r, 'C')).trim();
            if (!sigName) continue;

            // Skip header rows
            if (sigName === '信号名称' || sigName === 'Signal Name' || sigName === '信号列表') continue;

            // Replace spaces in signal names with underscores (LDF requires identifier tokens)
            sigName = sigName.replace(/\s+/g, '_');

            // Per-frame dedup on original name (before cross-frame renaming)
            const rawName = sigName;
            if (usedOriginalNamesInFrame.has(rawName)) {
                let suffix = 1;
                while (usedOriginalNamesInFrame.has(`${rawName}_${suffix}`)) suffix++;
                sigName = `${rawName}_${suffix}`;
            }
            usedOriginalNamesInFrame.add(sigName);

            // Parse bit range from column B (位 bit)
            const bitRangeStr = safeStr(sheet.getCell(r, 'B')).trim();
            let startBit = 0;
            let bitLen = 1;

            if (bitRangeStr) {
                // Format: "0~3", "4~6", "7", "8~10", "16~17"
                const rangeMatch = bitRangeStr.match(/(\d+)\s*[~\-–]\s*(\d+)/);
                if (rangeMatch) {
                    startBit = parseInt(rangeMatch[1]);
                    bitLen = parseInt(rangeMatch[2]) - startBit + 1;
                } else {
                    const singleMatch = bitRangeStr.match(/(\d+)/);
                    if (singleMatch) {
                        startBit = parseInt(singleMatch[1]);
                        bitLen = 1;
                    }
                }
            }

            // Factor and Offset
            const factorStr = safeStr(sheet.getCell(r, 'E')).trim();
            const offsetStr = safeStr(sheet.getCell(r, 'F')).trim();
            const factor = factorStr ? parseFloat(factorStr) : 1.0;
            const offset = offsetStr ? parseFloat(offsetStr) : 0;

            // Min/Max from columns J-K
            const minStr = safeStr(sheet.getCell(r, 'J')).trim();
            const maxStr = safeStr(sheet.getCell(r, 'K')).trim();
            const minPhys = minStr ? parseFloat(minStr) : 0;
            const maxPhys = maxStr ? parseFloat(maxStr) : Math.pow(2, bitLen) - 1;

            // Init value from column L (default value)
            const initStr = safeStr(sheet.getCell(r, 'L')).trim();
            // Unit from column D
            const unit = safeStr(sheet.getCell(r, 'D')).trim();

            // Value description from column O
            const valDescStr = safeStr(sheet.getCell(r, 'O')).trim();

            // Per-signal publisher/subscriber from M/N columns
            const sigPubRaw = safeStr(sheet.getCell(r, 'M')).trim();
            const sigSubRaw = safeStr(sheet.getCell(r, 'N')).trim();
            const sigPub = sigPubRaw ? extractNode(sigPubRaw) : pubName;
            const sigSub = sigSubRaw ? extractNode(sigSubRaw) : subName;

            // Cross-frame unique: if same signal name exists from another frame with a DIFFERENT publisher,
            // append _Publisher to disambiguate (LIN 2.0 requires globally unique signal names)
            const existingPublishers = globalSignalNames.get(sigName);
            if (existingPublishers) {
                if (!existingPublishers.has(sigPub)) {
                    existingPublishers.add(sigPub);
                    sigName = `${sigName}_${sigPub}`;
                }
            } else {
                globalSignalNames.set(sigName, new Set([sigPub]));
            }

            signals.push({
                name: sigName,
                startBit,
                bitLen,
                initValue: initStr || '0',
                factor,
                offset,
                minPhys,
                maxPhys,
                unit,
                valDesc: valDescStr,
                publisher: sigPub,
                subscriber: sigSub,
            });
        }

        // Determine if master or slave frame from schedule
        const isMasterFrame = scheduleNodes[si] === '主节点';
        // For the frame publisher in Frames section, LIN uses the node that provides the response
        // Master frames → publisher is the master (LSCM or BDC)
        // Slave frames → publisher is the slave

        frames.push({
            name: frameName,
            id: frameId,
            length: msgLength,
            publisher: pubName,
            subscriber: subName,
            signals,
            isMasterFrame,
            scheduleNodeType: scheduleNodes[si] || '从节点',
        });
    }

    // Collect all nodes
    const allNodes = new Set();
    for (const f of frames) {
        allNodes.add(f.publisher);
        allNodes.add(f.subscriber);
    }

    // Determine master and slaves
    // Master is typically LSCM (左侧智能座椅终端), slaves are the others
    const masterName = 'LSCM';  // This is the LIN master
    const slaveNames = [...allNodes].filter(n => n !== masterName);

    console.log(`主节点: ${masterName}`);
    console.log(`从节点: ${slaveNames.join(', ')}`);

    // Build schedule entries
    // Schedule table format: 主节点 frames first, then 从节点 frames
    // Each frame appears once per cycle with timeBase delay
    const scheduleDelay = timeBase;

    // Clean up temp
    rmSync(tmpDir, { recursive: true, force: true });

    return { frames, masterName, slaveNames, baudRate, scheduleDelay, scheduleCycle, timeBase };
}

// ─────────────────────────────────────────────────────────────
// LDF 生成
// ─────────────────────────────────────────────────────────────

function genLdf(data) {
    const { frames, masterName, slaveNames, baudRate, scheduleDelay } = data;
    const lines = [];

    // 1. Two empty lines + LIN_description_file
    lines.push('');
    lines.push('');
    lines.push('LIN_description_file;');
    lines.push('LIN_protocol_version = "2.0";');
    lines.push('LIN_language_version = "2.0";');
    lines.push(`LIN_speed = ${baudRate} kbps;`);
    lines.push('');

    // 2. Nodes
    lines.push('Nodes {');
    lines.push(`  Master: ${masterName}, 10 ms, 0.1 ms ;`);
    if (slaveNames.length > 0) {
        lines.push(`  Slaves: ${slaveNames.join(', ')} ;`);
    }
    lines.push('}');
    lines.push('');

    // 3. Signals (deduplicated by name — merge subscribers for same-named signals)
    const signalMap = new Map();
    for (const f of frames) {
        for (const sig of f.signals) {
            if (!signalMap.has(sig.name)) {
                signalMap.set(sig.name, {
                    bitLen: sig.bitLen,
                    initValue: sig.initValue,
                    publisher: sig.publisher,
                    subscribers: new Set([sig.subscriber]),
                });
            } else {
                signalMap.get(sig.name).subscribers.add(sig.subscriber);
            }
        }
    }
    lines.push('Signals {');
    for (const [name, sig] of signalMap) {
        let initVal = sig.initValue;
        if (initVal.startsWith('0x') || initVal.startsWith('0X')) {
            try { initVal = String(parseInt(initVal, 16)); } catch { initVal = '0'; }
        } else {
            const n = parseInt(initVal);
            if (!isNaN(n)) initVal = String(n);
            else initVal = '0';
        }
        const subs = [...sig.subscribers].join(', ');
        lines.push(`  ${name}: ${sig.bitLen}, ${initVal}, ${sig.publisher}, ${subs} ;`);
    }
    lines.push('}');
    lines.push('');

    // 4. Diagnostic_signals
    lines.push('Diagnostic_signals {');
    for (let i = 0; i < 8; i++) lines.push(`  MasterReqB${i}: 8, 0 ;`);
    for (let i = 0; i < 8; i++) lines.push(`  SlaveRespB${i}: 8, 0 ;`);
    lines.push('}');
    lines.push('');

    // 5. Frames
    lines.push('Frames {');
    for (const f of frames) {
        // Frame ID in decimal
        const fidDec = f.id;
        lines.push(`  ${f.name}: ${fidDec}, ${f.publisher}, ${f.length} {`);
        for (const sig of f.signals) {
            lines.push(`    ${sig.name}, ${sig.startBit} ;`);
        }
        lines.push('  }');
    }
    lines.push('}');
    lines.push('');

    // 6. Diagnostic_frames
    lines.push('Diagnostic_frames {');
    lines.push('  MasterReq: 0x3c {');
    for (let i = 0; i < 8; i++) lines.push(`    MasterReqB${i}, ${i * 8} ;`);
    lines.push('  }');
    lines.push('  SlaveResp: 0x3d {');
    for (let i = 0; i < 8; i++) lines.push(`    SlaveRespB${i}, ${i * 8} ;`);
    lines.push('  }');
    lines.push('}');
    lines.push('');

    // 7. Node_attributes
    lines.push('Node_attributes {');
    for (const slave of slaveNames) {
        // Assign NAD based on order
        const nadVal = slaveNames.indexOf(slave) + 1;
        lines.push(`  ${slave}{`);
        lines.push('    LIN_protocol = "2.0" ;');
        lines.push(`    configured_NAD = 0x${nadVal.toString(16).padStart(2, '0')} ;`);
        lines.push('    product_id = 0x0, 0x0, 255 ;');
        lines.push('    P2_min = 0 ms ;');
        lines.push('    ST_min = 0 ms ;');
        lines.push('    configurable_frames {');
        // Only include frames where the slave is the publisher
        for (const f of frames) {
            if (f.publisher === slave || f.subscriber === slave) {
                lines.push(`      ${f.name} = 0x0 ;`);
            }
        }
        lines.push('    }');
        lines.push('  }');
    }
    lines.push('}');
    lines.push('');

    // 8. Schedule_tables
    lines.push('Schedule_tables {');
    lines.push(` ${masterName}_Schedule {`);
    for (const f of frames) {
        lines.push(`    ${f.name} delay ${scheduleDelay} ms ;`);
    }
    lines.push('  }');
    lines.push('}');
    lines.push('');

    // 9. Signal_encoding_types (deduplicated by name)
    const emittedEncodings = new Set();
    const encodingBlocks = [];  // { name, lines[] }
    for (const f of frames) {
        for (const sig of f.signals) {
            if (emittedEncodings.has(sig.name)) continue;
            emittedEncodings.add(sig.name);

            const hasValDesc = sig.valDesc.length > 0;
            const hasPhysEncoding = (sig.factor !== 1.0 || sig.offset !== 0);
            if (!hasValDesc && !hasPhysEncoding) continue;

            const block = { name: sig.name, innerLines: [] };

            // logical_value from value descriptions
            if (hasValDesc) {
                const parsed = parseValDesc(sig.valDesc);
                for (const [val, desc] of parsed) {
                    if (desc) {
                        const d = desc.replace(/"/g, "'");
                        block.innerLines.push(`    logical_value, ${val}, "${d}" ;`);
                    } else {
                        block.innerLines.push(`    logical_value, ${val} ;`);
                    }
                }
            }

            // physical_value for non-default factor/offset
            if (hasPhysEncoding) {
                const maxRawVal = Math.pow(2, sig.bitLen) - 1;
                let minRaw = Math.round(sig.minPhys / sig.factor - sig.offset);
                let maxRaw = Math.round(sig.maxPhys / sig.factor - sig.offset);
                // Cap raw values within valid signal bit range
                if (minRaw < 0) minRaw = 0;
                if (maxRaw > maxRawVal) maxRaw = maxRawVal;
                if (maxRaw < minRaw) maxRaw = minRaw;
                block.innerLines.push(`    physical_value, ${minRaw}, ${maxRaw}, ${sig.factor}, ${sig.offset} ;`);

                // physical_value must be accompanied by logical_value
                if (!hasValDesc) {
                    block.innerLines.push(`    logical_value, ${minRaw}, "${sig.minPhys}" ;`);
                    block.innerLines.push(`    logical_value, ${maxRaw}, "${sig.maxPhys}" ;`);
                }
            }

            if (block.innerLines.length > 0) {
                encodingBlocks.push(block);
            }
        }
    }

    lines.push('Signal_encoding_types {');
    for (const block of encodingBlocks) {
        lines.push(`  ${block.name}_Encoding {`);
        for (const il of block.innerLines) lines.push(il);
        lines.push('  }');
    }
    lines.push('}');
    lines.push('');

    // 10. Signal_representation (deduplicated)
    const emittedRepr = new Set();
    lines.push('Signal_representation {');
    for (const f of frames) {
        for (const sig of f.signals) {
            if (emittedRepr.has(sig.name)) continue;
            const hasValDesc = sig.valDesc.length > 0;
            const hasPhysEncoding = (sig.factor !== 1.0 || sig.offset !== 0);
            if (hasValDesc || hasPhysEncoding) {
                // Only emit if there's a corresponding encoding block with content
                const encBlock = encodingBlocks.find(b => b.name === sig.name);
                if (encBlock) {
                    emittedRepr.add(sig.name);
                    lines.push(`  ${sig.name}_Encoding: ${sig.name} ;`);
                }
            }
        }
    }
    lines.push('}');
    lines.push('');

    return lines.join('\n');
}

// ─────────────────────────────────────────────────────────────
// 辅助函数
// ─────────────────────────────────────────────────────────────

function parseValDesc(text) {
    const result = [];
    if (!text) return result;
    // Replace HTML entity &#10; with actual newlines (fast-xml-parser preserves them)
    const normalized = text.replace(/&#10;/g, '\n').replace(/&#13;/g, '');
    for (const line of normalized.split('\n')) {
        const trimmed = line.trim();
        if (!trimmed) continue;

        // Handle range values: "0x2~0x3:Reserved"
        if (trimmed.includes('~') || trimmed.includes('～')) {
            const sep = trimmed.includes('～') ? '～' : '~';
            const parts = trimmed.split(sep, 2);
            const rangePart = parts[0].trim();
            const rest = parts[1];
            let desc = '';
            let endValStr = rest;

            // Find the colon after the end value
            const colonIdx = rest.indexOf(':');
            const fcolonIdx = rest.indexOf('：');
            const useColon = (fcolonIdx >= 0 && (colonIdx < 0 || fcolonIdx < colonIdx)) ? fcolonIdx : colonIdx;

            if (useColon >= 0) {
                endValStr = rest.substring(0, useColon).trim();
                desc = rest.substring(useColon + 1).trim();
            }

            try {
                const startVal = parseInt(rangePart, 16);
                const endVal = parseInt(endValStr, 16);
                if (!isNaN(startVal) && !isNaN(endVal)) {
                    // For range, we can't represent all values, use start/end as logical values
                    if (desc) {
                        result.push([startVal, `${desc}`]);
                        if (endVal !== startVal) {
                            result.push([endVal, `${desc}`]);
                        }
                    }
                }
            } catch { /* ignore */ }
            continue;
        }

        // Handle single value: "0x0:desc" or "0:desc"
        const colonIdx = trimmed.indexOf(':');
        const fcolonIdx = trimmed.indexOf('：');
        const useColon = (fcolonIdx >= 0 && (colonIdx < 0 || fcolonIdx < colonIdx)) ? fcolonIdx : colonIdx;

        if (useColon < 0) continue;

        const valStr = trimmed.substring(0, useColon).trim();
        let desc = trimmed.substring(useColon + 1).trim();

        try {
            let val;
            if (valStr.startsWith('0x') || valStr.startsWith('0X')) {
                val = parseInt(valStr, 16);
            } else if (/^\d+$/.test(valStr)) {
                val = parseInt(valStr, 10);
            } else {
                continue;
            }
            if (!isNaN(val)) {
                result.push([val, desc]);
            }
        } catch { /* ignore */ }
    }
    return result;
}

// ─────────────────────────────────────────────────────────────
// 保存文件
// ─────────────────────────────────────────────────────────────

function saveLdf(content, outputPath) {
    // UTF-8 without BOM + CRLF
    const crlfContent = content.replace(/\n/g, '\r\n');
    writeFileSync(outputPath, crlfContent, 'utf-8');
}

// ─────────────────────────────────────────────────────────────
// 验证
// ─────────────────────────────────────────────────────────────

function validateLdf(content) {
    const opens = (content.match(/{/g) || []).length;
    const closes = (content.match(/}/g) || []).length;
    const balanced = opens === closes;
    return { opens, closes, balanced };
}

// ─────────────────────────────────────────────────────────────
// 入口
// ─────────────────────────────────────────────────────────────

function main() {
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('用法: node lin_to_ldf.mjs <input.xlsx> [-o output.ldf]');
        process.exit(1);
    }

    const inputPath = args[0];
    let outputPath;
    const oIdx = args.indexOf('-o');
    if (oIdx >= 0 && oIdx + 1 < args.length) {
        outputPath = args[oIdx + 1];
    } else {
        const base = inputPath.replace(/\.xlsx$/i, '');
        outputPath = base + '.ldf';
    }

    console.log(`输入: ${inputPath}`);
    console.log(`输出: ${outputPath}\n`);

    const data = convert(inputPath);

    if (data.frames.length === 0) {
        console.log('错误: 没有解析到任何帧');
        process.exit(1);
    }

    const content = genLdf(data);
    saveLdf(content, outputPath);

    // Validation
    const { opens, closes, balanced } = validateLdf(content);
    const totalSignals = data.frames.reduce((sum, f) => sum + f.signals.length, 0);
    const totalEncodings = data.frames.reduce((sum, f) =>
        sum + f.signals.filter(s => s.valDesc || s.factor !== 1.0 || s.offset !== 0).length, 0);

    console.log(`\n[OK] LDF: ${outputPath}`);
    console.log(`   帧数: ${data.frames.length}`);
    console.log(`   信号数: ${totalSignals}`);
    console.log(`   编码数: ${totalEncodings}`);
    console.log(`   花括号: {=${opens}, }=${closes}, 平衡=${balanced}`);

    // List signals without encoding
    const noEnc = [];
    for (const f of data.frames) {
        for (const sig of f.signals) {
            if (!sig.valDesc && sig.factor === 1.0 && sig.offset === 0) {
                noEnc.push(sig.name);
            }
        }
    }
    if (noEnc.length > 0) {
        console.log(`   无编码: ${noEnc.join(', ')}`);
    }

    console.log('生成完成!');
}

main();
