using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace DromParser
{
    internal class ExcelCreater
    {
        /// <summary>
        /// Создаёт excel файл
        /// </summary>
        /// <param name="path">Имя сохраняемого файла, включая полный путь</param>
        /// <param name="dataTable"></param>
        /// <param name="templateExcel">Путь к шаблону или пустому excel</param>
        public void SaveExel(string path, DataTable dataTable, string listName, string templateExcel)
        {
            Excel.Application xlApp = new Excel.Application(); // запускаем Excel для чтения входного файла
            Excel.Workbook wb = xlApp.Workbooks.Open(templateExcel); // открываем шаблон указываю документ, в который я буду записывать данные
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[listName];   //  выбор листа 
            int x = 1; // строка листа excel
            foreach (DataRow nc in dataTable.Rows)
            {
                ws.Cells[x, 1] = nc[0].ToString(); //F1 - название столбца в dataTable dt  //ws.Cells[x, 1]  - x-строка в excel документе; 1 - столбец
                ws.Cells[x, 2] = nc[1].ToString(); //F2 - название столбца в dataTable dt 
                ws.Cells[x, 3] = nc[2].ToString();
                ws.Cells[x, 4] = nc[3].ToString();
                x++;
            }

            wb.SaveAs(path);
            wb.Close();
            xlApp.Quit();
        }
    }
}
