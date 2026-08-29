// Flattening a workbook on disk to text a model can read. Empty cells are kept
// as empty columns on purpose: on a quote sheet the meaning of a figure is the
// column it sits under, and collapsing the gaps silently shifts every value
// left — `Line Haul <tab> per diem excl` reads as a price in the second column.
import ExcelJS from 'exceljs';

/** Render every sheet of the workbook at `filePath` as tab-separated rows. */
export async function readXlsxText(filePath: string): Promise<string> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);

    const lines: string[] = [];
    wb.worksheets.forEach((ws) => {
        lines.push(`# ${ws.name}`);
        ws.eachRow({ includeEmpty: false }, (row) => {
            const cells: string[] = [];
            row.eachCell({ includeEmpty: true }, (cell) => {
                cells.push(String(cell.text ?? '').trim());
            });
            // Trailing blanks sit under no heading — a ragged tail of tabs is noise.
            while (cells.length && !cells[cells.length - 1]) cells.pop();
            if (cells.some(Boolean)) lines.push(cells.join('\t'));
        });
    });
    return lines.join('\n');
}
