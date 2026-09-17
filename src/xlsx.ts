// Flattening a workbook on disk to text a model can read. Empty cells stay as
// empty columns on purpose: a figure's meaning is often the column it sits
// under, and collapsing the gaps would silently shift every later value left.
import ExcelJS from 'exceljs';

/** Every sheet of the workbook at `filePath`, as a `# name` heading and tab-separated rows. */
export async function readXlsxText(filePath: string): Promise<string> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const lines: string[] = [];
    workbook.worksheets.forEach(sheet => {
        lines.push(`# ${sheet.name}`);
        sheet.eachRow({ includeEmpty: false }, row => {
            const cells: string[] = [];
            row.eachCell({ includeEmpty: true }, cell => {
                cells.push(String(cell.text ?? '').trim());
            });
            // Trailing blanks sit under no heading; a ragged tail of tabs is noise.
            while (cells.length && !cells[cells.length - 1]) cells.pop();
            if (cells.some(Boolean)) lines.push(cells.join('\t'));
        });
    });
    return lines.join('\n');
}
