import ExcelJS from 'exceljs';

type ExportRow = Record<string, unknown>;

const EXCEL_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const CSV_MIME_TYPE = 'text/csv;charset=utf-8';

export async function exportToExcel(
  data: ExportRow[],
  columns: { key: string; header: string }[],
  filename: string,
): Promise<void> {
  const isCsv = filename.toLowerCase().endsWith('.csv');
  const extension = isCsv ? 'csv' : 'xlsx';
  const downloadFilename = filename.toLowerCase().endsWith(`.${extension}`)
    ? filename
    : `${filename}.${extension}`;

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet1');

  worksheet.columns = columns.map(({ key, header }) => ({ key, header }));
  data.forEach((row) => {
    worksheet.addRow(
      Object.fromEntries(columns.map(({ key }) => [key, row[key] ?? ''])),
    );
  });

  const buffer = isCsv
    ? await workbook.csv.writeBuffer()
    : await workbook.xlsx.writeBuffer();
  const blob = new Blob([new Uint8Array(buffer)], {
    type: isCsv ? CSV_MIME_TYPE : EXCEL_MIME_TYPE,
  });
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = downloadUrl;
  link.download = downloadFilename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(downloadUrl);
}
