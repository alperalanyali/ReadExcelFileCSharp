using System;
using System.Text;
using GemBox.Spreadsheet;
namespace ReadExcelFile
{
    public class ReadExcelFileClass
    {
       public static string ReadExcelFile()
        {
            string sonuc = "NOK";

            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                var workbook = ExcelFile.Load("/Users/alperalanyali/Desktop/vardiyaplani2.xlsx");
                var sb = new StringBuilder();



                // Iterate through all worksheets 
                foreach (var worksheet in workbook.Worksheets)
                {

                    sb.AppendLine();
                    sb.AppendFormat("{0} {1} {0}", new string('-', 25), worksheet.Name);
                    CellRange colCount = worksheet.Columns[1].Cells;
                    int rowCount = worksheet.Rows.Count;
                    Console.WriteLine("{0} {1}",rowCount ,colCount);
                    // Iterate through all rows in an Excel worksheet.
                    foreach (var row in worksheet.Rows)
                    {
                        Console.WriteLine(row.AllocatedCells.ToString());
                        sb.AppendLine("");
                        //sb.Append(r);
                        // Iterate through all allocated cells in an Excel row.

                        foreach (var cell in row.AllocatedCells)
                            if (cell.ValueType != CellValueType.Null)
                                sb.Append(string.Format("{0}", cell.Value).PadRight(25));
                            else
                                sb.Append(new string(' ', 25));
                    }
                }

          //      Console.WriteLine(sb.ToString());
            }
            catch(Exception ex)
            {
                sonuc += " Ex Mesaage:" + ex.Message;
            }




            return sonuc;
        }
    }
}
