using System;

namespace DromParser
{
    internal class Program
    {
        static Parser parser { get; set; }
        static void Main(string[] args)
        {
            parser = new Parser();
            parser.ready += Parser_ready;
            parser.Start("lada", "granta", 10);
        }

        private static void Parser_ready()
        {
            Console.WriteLine("OK");

            ExcelCreater excelCreater = new ExcelCreater();

            excelCreater.SaveExel(System.IO.Directory.GetCurrentDirectory() + "\\result.xlsx",
                parser.DataTable, "Лист1", System.IO.Directory.GetCurrentDirectory() + "\\template.xlsx");
        }
    }
}
