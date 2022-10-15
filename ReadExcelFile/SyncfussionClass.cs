using System;
using System.IO;
using Syncfusion.XlsIO;

namespace ReadExcelFile
{
    public class SyncfussionClass
    {
        public static void Excel()
        {
            ExcelEngine excelEngine = new ExcelEngine();

            IApplication application = excelEngine.Excel;

            application.DefaultVersion = ExcelVersion.Excel2010;

            string basePath = "/Users/alperalanyali/Desktop/vardiyaplani2.xlsx";
            FileStream sampleFile = new FileStream(basePath, FileMode.Open);

            IWorkbook workbook = application.Workbooks.Open(sampleFile);
            IWorksheet workSheet = workbook.Worksheets[0];
            var rowCount = workSheet.UsedRange.Row;
            var colCount = workSheet.UsedRange.Column;
            Console.WriteLine("RowCount:{0}\tColumnCo:{1}",rowCount, colCount);
        }
    }
}
