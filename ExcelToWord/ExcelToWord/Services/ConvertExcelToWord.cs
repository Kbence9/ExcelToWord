using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelToWord.Services
{
    static class ConvertExcelToWord
    {
        public static void ConvertFile(string excel, string word)
        {

            Microsoft.Office.Interop.Excel.Application _excelApp = new()
            {
                Visible = true
            };

            string fileName = excel;

            Console.WriteLine(excel);

            try
            {
                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Microsoft.Office.Interop.Excel.Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);

                //access the cells
                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        //access each cell
                        Debug.Print(valueArray[row, col].ToString());
                    }
                }

                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);

                _excelApp.Quit();
                Marshal.FinalReleaseComObject(_excelApp);
            } 
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }
    }
}
