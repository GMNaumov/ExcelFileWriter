using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelFileWriter
{
    internal class ExcelDataWriter : DataWriter
    {
        private Excel.Application excelApplication;
        private readonly object missValue;

        public ExcelDataWriter()
        {
            missValue = System.Reflection.Missing.Value;
        }

        public void Write(List<object> objects)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Запись в ячейки листа Excel производится присваиванием значению Cell[строка, столбец]
        /// Нумерация значений начинается со значения 1.
        /// </summary>
        /// <param name="cars"></param>
        public void Write(List<Car> cars, string path)
        {
            excelApplication = new Application();
            Workbook excelWorkbook = excelApplication.Workbooks.Add();
            Worksheet excelWorksheet = excelWorkbook.Worksheets[1];

            try
            {
                SetTableHeaders(excelWorksheet);
                SetHeaderStyle(excelWorksheet);
                
                for (int i = 0; i < cars.Count; i++)
                {
                    Car next = cars[i];

                    int cellNum = i + 2;

                    excelWorksheet.Cells[cellNum, 1] = next.Brand;
                    excelWorksheet.Cells[cellNum, 2] = next.Model;
                    excelWorksheet.Cells[cellNum, 3] = next.Price;
                }


                excelApplication.Visible = false;
                excelWorkbook.SaveAs2(path);
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception has raised!");
                Console.WriteLine(exception.StackTrace);
            }
            finally
            {
                excelWorkbook.Close();
                excelApplication.Quit();
                Marshal.ReleaseComObject(excelWorksheet);
                Marshal.ReleaseComObject(excelWorkbook);
                Marshal.ReleaseComObject(excelApplication);
                GC.Collect();
            }
        }

        /// <summary>
        /// Устанавливаем заголовки таблицы
        /// </summary>
        /// <param name="worksheet"></param>
        private void SetTableHeaders(Worksheet worksheet)
        {
            worksheet.Range["A1"].Value = "Car Brand";
            worksheet.Range["B1"].Value = "Car Model";
            worksheet.Range["C1"].Value = "Car Price";
        }

        private void SetHeaderStyle(Worksheet worksheet)
        {
            worksheet.get_Range("A1", "C1").Style.Font.Size = 20;
        }
    }
}
