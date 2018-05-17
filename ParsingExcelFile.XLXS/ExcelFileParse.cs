using System;
using Microsoft.Office.Interop.Excel;

using ParsingExcelFile.Consts;

namespace ParsingExcelFile.XLSX
{
    public class ExcelFileParse
    {
        private Application xlApp;
        private Workbook xlWorkbook;
        private Worksheet xlWorksheet;
        private Range xlRange;
        
        public ExcelFileParse(string path)
        {
            xlApp = new Application();
            xlWorkbook = xlApp.Workbooks.Open(path);
            xlWorksheet = (Worksheet) xlWorkbook.Worksheets.get_Item(1);
        }
        static void Main(string[] args)
        {
            string prefix = @"e:\\demo-excel-workbooks\\";
            ExcelFileParse xlFileParse = new ExcelFileParse(prefix + FilePaths.paths[0]);
            xlFileParse.xlRange = xlFileParse.xlWorksheet.UsedRange;

            try
            {
                for (var r = 1; r <= xlFileParse.xlRange.Rows.Count; r++)
                {
                    for (var c = 1; c <= xlFileParse.xlRange.Columns.Count; c++)
                    {
                        Console.Write((xlFileParse.xlRange.Cells[r, c]).Value + "\t");
                    }
                    Console.WriteLine();
                }
                Console.WriteLine($"\nTotal Records: {xlFileParse.xlRange.Rows.Count}");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                xlFileParse.xlWorkbook.Close();
            }

            Console.ReadLine();
        }
    }
}
