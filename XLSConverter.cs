using OfficeOpenXml;
using OfficeOpenXml.Style;

class XLSConverter
{
    private string? outFile;
    private int iRow = 1;
    private int iCol = 1;
    private ExcelPackage package;
    private ExcelWorksheet worksheet;

    public XLSConverter(string outFile)
    {
        package = new ExcelPackage(new FileInfo(outFile));
        worksheet = package.Workbook.Worksheets.Add("Sheet1");
    }

    public void Convert(IEnumerable<string[]> lines)
    {
        foreach (string[] line in lines)
        {
            foreach (string col in line)
            {
                worksheet.Cells[iRow, iCol].Value = col;
                if (iRow == 1)
                {
                    highlightColumn();
                }
                iCol++;
            }
            iRow++;
            iCol = 1;
        }

        // Freeze the top row
        worksheet.View.FreezePanes(2, 1); // 2 is the row below the top row, 1 is the first column

        package.Save();
    }

    private void highlightColumn()
    {
        worksheet.Cells[iRow, iCol].Style.Font.Bold = true;
        worksheet.Cells[iRow, iCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[iRow, iCol].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
    }
}