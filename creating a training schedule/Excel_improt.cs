using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_improt
{
    class Excel_improt : IDisposable
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private string _filePath;

        public Excel_improt()
        {
            _excel = new Excel.Application();
        }

        public void Quit()
        {
            try
            {
                Process[] localByName = Process.GetProcessesByName("Excel");
                foreach (Process p in localByName)
                {
                    p.Kill();
                }
            }
            catch { }

        }


        internal void delete(string filePath)
        {

            if (File.Exists(filePath))
            {
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception e)
                {
                }
            }
            else
            {
            }
        }

        public void print_doc(string printerName) // параметр принтера
        {
            try
            {
                Excel._Worksheet xlWorksheet = _workbook.ActiveSheet;

                xlWorksheet.PageSetup.Zoom = false;
                //if (lenght >= 3)
                xlWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                xlWorksheet.PageSetup.FitToPagesWide = 1;
                xlWorksheet.PageSetup.FitToPagesTall = 1;

                xlWorksheet.PageSetup.LeftMargin = _excel.InchesToPoints(0);
                xlWorksheet.PageSetup.RightMargin = _excel.InchesToPoints(0);
                xlWorksheet.PageSetup.TopMargin = _excel.InchesToPoints(0);
                xlWorksheet.PageSetup.BottomMargin = _excel.InchesToPoints(0);

                xlWorksheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing, printerName, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }


        public void format_cells(string name_podpis)
        {
            try
            {
                Excel._Worksheet xlWorksheet = _workbook.ActiveSheet;

                string work_cells = "B1:O23";
                xlWorksheet.Range["B2:O23"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                xlWorksheet.Range[work_cells].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorksheet.Range[work_cells].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //выставление ширины столбцов
                xlWorksheet.Range[work_cells].ColumnWidth = 14;


                string[] temp_range = { "B3:B5", "B6:B8", "B9:B11", "B12:B14", "B15:B17", "B18:B20", "B21:B23" };
                Range range;

                //объединение ячеек
                for (int i = 0; i < temp_range.Length; i++)
                {
                    range = xlWorksheet.Range[temp_range[i]];
                    range.Merge();
                }

                //перенос слов
                xlWorksheet.Range["B2:O23"].EntireColumn.WrapText = true;


                //подпись документа
                string podpis = "B1:O1";
                xlWorksheet.Cells[1, 1].EntireRow.Font.Bold = true;
                range = xlWorksheet.Range[podpis];
                range.Font.Size = 16;
                range.Merge();
                range.Value = "Дворец спорта Кузнецких металлургов";
                
                podpis = "B2:O2";
                range = xlWorksheet.Range[podpis];
                range.Merge();
                range.Value = name_podpis;

                podpis = "C4:C4";
                range = xlWorksheet.Range[podpis];

                for (char c = 'C'; c <= 'O'; c++)
                {
                    for (int i = 3; i < 24; i++)
                    {
                        string a;
                        podpis = Convert.ToString(c.ToString() + i.ToString() + ":" + c.ToString() + i.ToString());
                        range = xlWorksheet.Range[podpis];
                        a = Convert.ToString(range.Value);

                        if (a == "-")
                            xlWorksheet.Range[podpis].Interior.ColorIndex = 56;
                    }
                }
            }
            catch (Exception ex) { /*MessageBox.Show(ex.Message);*/ }
        }
        internal bool Open(string filePath)
        {

            try
            {
                _workbook = _excel.Workbooks.Add();

                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }

                //выбор 1 книги
                Excel._Worksheet xlWorksheet = _workbook.Sheets[1];

                xlWorksheet = _workbook.ActiveSheet;
                //установка длины для шрифта
                Excel.Range xlRange = xlWorksheet.Range["A1", "Z90"];

                //xlRange.ColumnWidth = 30;
                xlRange.Font.Name = "Times New Roman";
                xlRange.Font.Size = 14;

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }

        internal bool Set(string column, int row, string data)
        {
            try
            {
                // var val = ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;

                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { /*MessageBox.Show(ex.Message); */}
            return false;
        }

        internal object Get(string column, int row)
        {
            try
            {
                return ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return null;
        }

        public void Dispose()
        {
            try
            {
                //_workbook.Close();
                //_excel.Quit();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
