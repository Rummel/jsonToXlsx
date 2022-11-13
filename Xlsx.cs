
using ClosedXML.Excel;

public class Xlsx
{
    XlsxData data;
    public Xlsx(XlsxData data)
    {
        this.data = data;
    }

    private Boolean checkConditions()
    {
        if (this.data.FileNameTarget == null)
        {
            Console.WriteLine("no target file");
            return false;
        }

        if (this.data.Worksheets == null)
        {
            Console.WriteLine("no worksheets to convert");
            return false;
        }
        return true;
    }


    public void Convert()
    {
        //string? fileNameSource = string.IsNullOrEmpty(data.fileNameSource) ? data.fileNameTarget : data.fileNameSource;
        if (!this.checkConditions() || (this.data.Worksheets == null))
        {
            return;
        }

        string? fileNameSource = data.FileNameSource;
        using (var workbook = File.Exists(fileNameSource) ? new XLWorkbook(fileNameSource) : new XLWorkbook())
        {
            foreach (var wbData in this.data.Worksheets)
            {
                var worksheetName = wbData.name;
                IXLWorksheet worksheet;
                if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out worksheet))
                {
                    worksheet = workbook.Worksheets.Add(worksheetName);
                }
                //worksheet.Cell("A1").Value = "Hello World!";
                //worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                //worksheet.Cells("C2:C4").Style.Fill.BackgroundColor = XLColor.Blue;
                foreach (var cellData in wbData.cells)
                {
                    if (cellData.Value != null)
                    {
                        if (cellData.Value is XlsxTableRow data)
                        {
                            worksheet.Cell(cellData.Cell).InsertData(data);
                        }
                        else
                        {
                            worksheet.Cells(cellData.Cell).Value = cellData.Value;
                        }
                    }

                    if (cellData.DataType != null)
                    {
                        worksheet.Cells(cellData.Cell).DataType = cellData.DataType.Value;
                    }
                    if (cellData.FormulaA1 != null)
                    {
                        worksheet.Cells(cellData.Cell).FormulaA1 = cellData.FormulaA1;
                    }
                    if (cellData.Style != null)
                    {
                        var style = cellData.Style;
                        if (style.Color.HasValue)
                        {
                            worksheet.Cells(cellData.Cell).Style.Fill.BackgroundColor = XLColor.FromArgb(style.Color.Value);
                        }
                        if (style.NumberFormat != null)
                        {
                            var format = style.NumberFormat;
                            if (format.Id.HasValue)
                            {
                                worksheet.Cells(cellData.Cell).Style.NumberFormat.NumberFormatId = format.Id.Value;
                            }
                            if (format.Format != null)
                            {
                                worksheet.Cells(cellData.Cell).Style.NumberFormat.Format = format.Format;
                            }
                        }

                        if (style.DateFormat != null)
                        {
                            var format = style.DateFormat;
                            if (format.Id.HasValue)
                            {
                                worksheet.Cells(cellData.Cell).Style.DateFormat.NumberFormatId = format.Id.Value;
                            }
                            if (format.Format != null)
                            {
                                worksheet.Cells(cellData.Cell).Style.DateFormat.Format = format.Format;
                            }
                        }
                    }
                }

            }
            Console.WriteLine("save file...");
            workbook.SaveAs(data.FileNameTarget);
        }

    }
}

