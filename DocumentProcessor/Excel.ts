

import * as XLSX from 'xlsx';

export interface ExcelTable {
    fileName: string;
    sheetName: string;
    data: Record<string, unknown>[];
}

export function parseExcelDynamic(data: ArrayBuffer, fileName: string): { fileName: string; tables: ExcelTable[] } {
    const workbook = XLSX.read(data, { type: 'array' });
    const allTables: ExcelTable[] = [];

    for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const tables = findTablesInSheet(worksheet, sheetName, fileName);
        allTables.push(...tables);
    }

    
    return { fileName, tables: allTables };
}

function findTablesInSheet(
    worksheet: XLSX.WorkSheet,
    sheetName: string,
    fileName: string
): ExcelTable[] {
    const tables: ExcelTable[] = [];
    const visited = new Set<string>();
    const range = XLSX.utils.decode_range(worksheet['!ref'] ?? 'A1');

    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (visited.has(cellAddress)) continue;

            const cell = worksheet[cellAddress] as XLSX.CellObject | undefined;
            if (cell?.v !== undefined && cell?.v !== null) {
                const tableRange = findTableRange(worksheet, R, C, visited);
                const data = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { range: tableRange, defval: null });
                if (data.length > 0) {
                    tables.push({
                        fileName: fileName,
                        sheetName: sheetName,
                        data: data
                    });
                }
            }
        }
    }

    return tables;
}

function findTableRange(
    worksheet: XLSX.WorkSheet,
    startR: number,
    startC: number,
    visited: Set<string>
): string {
    const range = XLSX.utils.decode_range(worksheet['!ref'] ?? 'A1');
    const queue = [{ r: startR, c: startC }];
    // const tableCells = new Set<string>();
    let minR = startR, maxR = startR, minC = startC, maxC = startC;

    while (queue.length > 0) {
        const { r, c } = queue.shift()!;
        const cellAddress = XLSX.utils.encode_cell({ r, c });

        if (
            r < range.s.r ||
            r > range.e.r ||
            c < range.s.c ||
            c > range.e.c ||
            visited.has(cellAddress)
        ) {
            continue;
        }

        const cell = worksheet[cellAddress] as XLSX.CellObject | undefined;
        if (cell?.v !== undefined && cell?.v !== null) {
            visited.add(cellAddress);
            // tableCells.add(cellAddress);

            minR = Math.min(minR, r);
            maxR = Math.max(maxR, r);
            minC = Math.min(minC, c);
            maxC = Math.max(maxC, c);

            // Check neighbors
            for (let dr = -1; dr <= 1; dr++) {
                for (let dc = -1; dc <= 1; dc++) {
                    if (dr === 0 && dc === 0) continue;
                    queue.push({ r: r + dr, c: c + dc });
                }
            }
        }
    }

    return XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
}


export function parseExcelFixed(
    data: ArrayBuffer,
    range: string,
    fileName: string,
    autoHeader: boolean
): { fileName: string; table?: ExcelTable; error?: string; message?: string } {
    const workbook = XLSX.read(data, { type: 'array' });

    const [sheetName, rangeStr] = range.split('!');
    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet) {
        return { error: `Sheet "${sheetName}" not found.`, fileName };
    }

    let dataArr: Record<string, unknown>[] = [];
    if (autoHeader) {
        // Get the range boundaries
        const ref = rangeStr || worksheet['!ref'];
        const range = XLSX.utils.decode_range(ref ?? 'A1');
        const colCount = range.e.c - range.s.c + 1;
        // Generate headers: ["Column_1", "Column_2", ...]
        const headers = Array.from({ length: colCount }, (_, i) => `Column_${i+1}`);
        dataArr = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
            range: rangeStr,
            defval: null,
            header: headers,
        });
    } else {
        dataArr = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, { range: rangeStr, defval: null });
    }

    if (dataArr.length === 0) {
        return { message: "No data found in the specified range.", fileName };
    }

    const result: ExcelTable = {
        fileName: fileName,
        sheetName: sheetName,
        data: dataArr
    };

    // Return as object
    return { fileName, table: result };
}

// Fixed column search: get all rows below a header row range until a blank row is hit in that range
export function parseExcelFixedColumnSearch(
    data: ArrayBuffer,
    range: string,
    fileName: string,
    autoHeader: boolean
): { fileName: string; table?: ExcelTable; error?: string; message?: string } {
    const workbook = XLSX.read(data, { type: 'array' });
    const [sheetName, rangeStr] = range.split('!');
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        return { error: `Sheet "${sheetName}" not found.`, fileName };
    }
    // Parse the range, e.g. B2:Z2
    const headerRange = XLSX.utils.decode_range(rangeStr);
    let headers: string[] = [];
    const colCount = headerRange.e.c - headerRange.s.c + 1;
    if (autoHeader) {
        headers = Array.from({ length: colCount }, (_, i) => `Column_${i + 1}`);
    } else {
        // Get header values from the specified row
        for (let c = headerRange.s.c; c <= headerRange.e.c; ++c) {
            const cellAddress = XLSX.utils.encode_cell({ r: headerRange.s.r, c });
            const cell = worksheet[cellAddress] as XLSX.CellObject | undefined;
            headers.push(cell?.v != null ? String(cell.v) : `Column_${c - headerRange.s.c + 1}`);
        }
    }
    // Now, collect rows below header until a blank row is hit in the range
    const rows: Record<string, unknown>[] = [];
    let row = autoHeader ? headerRange.s.r : headerRange.s.r + 1;
    while (true) {
        let isBlankRow = true;
        const rowObj: Record<string, unknown> = {};
        for (let c = headerRange.s.c; c <= headerRange.e.c; ++c) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c });
            const cell = worksheet[cellAddress] as XLSX.CellObject | undefined;
            const header = headers[c - headerRange.s.c];
            rowObj[header] = cell?.v ?? null;
            if (cell?.v !== undefined && cell?.v !== null && String(cell.v).trim() !== "") {
                isBlankRow = false;
            }
        }
        if (isBlankRow) break;
        rows.push(rowObj);
        row++;
    }
    if (rows.length === 0) {
        return { message: "No data found in the specified range.", fileName };
    }
    const result: ExcelTable = {
        fileName: fileName,
        sheetName: sheetName,
        data: rows
    };
    return { fileName, table: result };
}