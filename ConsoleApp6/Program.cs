using OfficeOpenXml;
using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.Style;
using System.Drawing;
using HtmlAgilityPack;
using System.Text;

namespace ConsoleApp1
{
    
    class ListCreator
    {
        private string Name;
        public List<string> Container;

        public ListCreator (string _name)
        {
            this.Name = _name;
            this.Container = new List<string> ();
        }
        public string PropertyName
        {
            get
            {
                return Name;
            }

            
        }
    }
    
    
    class Program
    {
        static void Main(string[] args)
        {

            string test = "http://images.firma-gamma.ru/images/4/8/g52096192552.jpg,http://images.firma-gamma.ru/images/f/a/g52096192552u.jpg";
            int index1 = test.LastIndexOf("http");
            Console.WriteLine(index1);
            Console.WriteLine(test.Substring(index1));
            string bombom = test.Remove(index1-1);
            Console.WriteLine(bombom);
            int index2 = bombom.IndexOf("/g");
            Console.WriteLine(index2);
            Console.WriteLine(bombom.Remove(0, index2+2).Replace(".jpg", ""));
            string test2 = "flag://images.firma-gamma.ru/images/d/5/g31128389392.jpg,http://images.firma-gamma.ru/images/5/d/g31128389392u.jpg,http://images.firma-gamma.ru/images/f/e/g31128389392u_1.jpg";
            
            
             
             int index3 = test2.IndexOf("http");
             Console.WriteLine(index3);
            string proba = test2.Substring(index3-1);
             Console.WriteLine(proba);
            string bombom2 = test2.Replace(proba, "").Replace("flag", "http");
            Console.WriteLine(bombom2);

            //План если значение не null  то начало http меняем на flag, если null, то записываем в ячейку "pkpkpk"
            //делаем список с zeroindex с индексом 0
            //добавляем в список все значения из столбика (индекс списка будет совпадать с номером строки)
            //в списке строки которые либо не содержат "http" либо содержать хотя бы одно "http" либо содержат "pkpkpk"
            //если элемент списка сожержит "http" и не равняется "pkpkpk", то ищем индекс ПЕРВОГО  вхождения и получаем строку с ПЕРВОГО-1 ИНДЕКСА
            //для этого же элемента полученную выше строку заменяем на "" и "flag" на "http"
            var html = @"http://shop.firma-gamma.ru/opengood/2465297952/#good2465297952";
            HtmlWeb web = new HtmlWeb();
            web.AutoDetectEncoding = false;
            web.OverrideEncoding = Encoding.GetEncoding ("windows-1251");
           

            var htmlDoc = web.Load(html);
            string Xpath = "//*[@id='good2465297952']/div[3]/div[1]/p";

            //Console.WriteLine("Look here                " + htmlDoc.GetElementbyId("good2465297952").WriteTo());

            var node = htmlDoc.DocumentNode.SelectSingleNode("/ html / head / title");
            
            Console.WriteLine(node.InnerText);
            
           // на гамме кодировка!!


            //Создание источника данных
            string MyFile = @"D:\c# excel\source.xlsx";
            FileInfo existingFile = new FileInfo(MyFile);
            ExcelPackage package = new ExcelPackage(existingFile);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            //Создание файла с которм мы будем работать
            string NewFile = @"D:\c# excel\target2.xlsx";
            FileInfo existingTargetFile = new FileInfo(NewFile);
            ExcelPackage Targetpackage = new ExcelPackage(existingTargetFile);
            // Создаем лист, в котором будем работать
            Targetpackage.Workbook.Worksheets.Add("MySheet");
            ExcelWorksheet Targetworksheet = Targetpackage.Workbook.Worksheets["MySheet"];
 
            List<string> books = new List<string>();
            string s = "nsnsns";
            string ExtractedIndexes;

            worksheet.Cells["A1:A80"].Copy(Targetworksheet.Cells["A1:A80"]);  // было 80 000
            worksheet.Cells["B1:B80"].Copy(Targetworksheet.Cells["B1:B80"]); // юыло 80 000
            // в таргетшит  сравниваем пред ячейку со след ячейкой и когда разные значения --- стваим другой цвет
            Targetworksheet.Cells[1, 1].Value = "nsnsns1";
            for (int i = 2; i < 80; i++) // было 80 000
            {
                if (Targetworksheet.Cells[i + 1, 1].Value != null && Targetworksheet.Cells[i, 1].Value.ToString() != Targetworksheet.Cells[i + 1, 1].Value.ToString())
                {
                    Targetworksheet.Cells[i + 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Targetworksheet.Cells[i + 1, 1].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                }
            }
            // пробегаемся по всей колонке в таргетшит --- если цветной фон то заменяем значение на "nsnsnsINDEX"
            for (int i = 1; i < 80; i++) // было 80 000
            {
                if (Targetworksheet.Cells[i, 1].Value != null && Targetworksheet.Cells[i, 1].Style.Fill.BackgroundColor.Rgb == "FFFFA500")
                {
                    Targetworksheet.Cells[i, 1].Value = s + i; // раньше было ns+1
                }
            }
            // пробегаемся по всей колонке А в таргетшит --- если значение содержит "nsINDEX"  дoбaвляем в список books
            for (int i = 1; i < 80; i++) // было 80 000
            {
                if (Targetworksheet.Cells[i, 1].Value != null && Targetworksheet.Cells[i, 1].Value.ToString() != "" && Targetworksheet.Cells[i, 1].Value.ToString().Contains(s)) // раньше было ns
                {
                    books.Add(Targetworksheet.Cells[i, 1].Value.ToString());
                }
            }
            List<string> BooksWithoutNs = new List<string>();
            // в букс убираем нс и получаем список индексов
            foreach (string Value in books)
            {
                ExtractedIndexes = Value.Remove(0, 6);
                BooksWithoutNs.Add(ExtractedIndexes);
            }
           
            
            List<string> KeyFromA = new List<string>();            
            // извлекаем по индексам значения в новый список 
            foreach (string Key in BooksWithoutNs)
            {
                Console.WriteLine(Key);
                KeyFromA.Add(worksheet.Cells[Int32.Parse(Key), 1].Value.ToString());
            }
            // делаем словарь  string "Молния....." (из списка KeyFromA) ------ класс.СПИСОК (это первый цикл), пробешаемся по колонке А из источника и если значение в ячейке равно слову из
            //списка KeyFromA то мы копируем в класс.СПИСОК значения из колонки B
            Dictionary<string, ListCreator> SupraList = new Dictionary <string, ListCreator> ();
            Dictionary<string, ListCreator> PriceList = new Dictionary<string, ListCreator>();
            Console.WriteLine("must be 50     " +package.Workbook.Worksheets[1].Cells[2, 4].Value);
            Console.WriteLine("must be 960    " + package.Workbook.Worksheets[1].Cells[2, 5].Value);
            Console.WriteLine(Convert.ToDouble(package.Workbook.Worksheets[1].Cells[13, 5].Value));
            double[] DevidedValues = new double[90];
            for (int i = 1; i < 80; i++)
            {
            
            DevidedValues[i-1] = Convert.ToDouble(package.Workbook.Worksheets[1].Cells[i, 5].Value.ToString()) / Convert.ToDouble(package.Workbook.Worksheets[1].Cells[i, 4].Value.ToString());
            Console.WriteLine("Test deviding    " + DevidedValues[i-1]);
            package.Workbook.Worksheets[1].Cells[i, 6].Value = DevidedValues[i - 1].ToString();
            }

            foreach (string Key in KeyFromA)
            {
                SupraList.Add(Key, new ListCreator (Key));
                PriceList.Add(Key, new ListCreator(Key));
                for (int i = 1; i < 80; i++) // было 80 000
                {
                if (package.Workbook.Worksheets[1].Cells[i, 1].Value != null && Key == package.Workbook.Worksheets[1].Cells[i, 1].Value.ToString() && package.Workbook.Worksheets[1].Cells[i, 2].Value != null) // в колонке B  есть пустые значения
                    {
                        SupraList[Key].Container.Add(package.Workbook.Worksheets[1].Cells[i, 2].Value.ToString());
                    }
                if (package.Workbook.Worksheets[1].Cells[i, 1].Value != null && Key == package.Workbook.Worksheets[1].Cells[i, 1].Value.ToString() && package.Workbook.Worksheets[1].Cells[i, 6].Value != null)
                    {
                        PriceList[Key].Container.Add(package.Workbook.Worksheets[1].Cells[i, 6].Value.ToString());
                    }
                }
            }
              // самое главное
            for (int i = 0; i < BooksWithoutNs.Count; i++)
            {
                Targetpackage.Workbook.Worksheets.Add(BooksWithoutNs[i]);
                Targetpackage.Workbook.Worksheets[BooksWithoutNs[i]].Cells[1, 1].Value = KeyFromA[i];
                
                for (int j = 0; j < SupraList[KeyFromA[i]].Container.Count; j++)
                {
                    Targetpackage.Workbook.Worksheets[BooksWithoutNs[i]].Cells[j + 2, 1].Value = SupraList[KeyFromA[i]].Container[j];

                }

                for (int k = 0; k < PriceList[KeyFromA[i]].Container.Count; k++)
                {
                    Targetpackage.Workbook.Worksheets[BooksWithoutNs[i]].Cells[k + 2, 2].Value = PriceList[KeyFromA[i]].Container[k]; // по идеи цена записывается в колонку b
                }

                Targetpackage.Workbook.Worksheets[BooksWithoutNs[i]].Cells[1, 2].Value = Targetpackage.Workbook.Worksheets[BooksWithoutNs[i]].Cells[2, 2].Value;
            }

            

            // Итого есть список с индексами-цифрами, список с ключевыми значениями и 
            //Словарь (Key - ключевое значение, Value - Класс.Список {значения копируються из второго столбика источника по ключевому слову}))
            // поставить ограничение = либо Count списка/слловаря или  последненму значение списка индексов или сделать из колонки А список и ограничить его Count*ом
            Targetpackage.Workbook.Worksheets.Delete("MySheet");
            List <string> ValueFrom = new List<string> ();
            List<string> PriceFrom = new List<string>();
            for (int i = 1; i < Targetpackage.Workbook.Worksheets.Count; i++)
            {
                Console.WriteLine(Targetpackage.Workbook.Worksheets[i]);

                for (int j = 1; j < 100; j++) // было 1000
                {
                    if (Targetpackage.Workbook.Worksheets[i].Cells[j, 1].Value != null)
                    {
                        ValueFrom.Add(Targetpackage.Workbook.Worksheets[i].Cells[j, 1].Value.ToString());
                    }
                }
                for (int k = 1; k < 100; k++) // было 1000
                {
                    if (Targetpackage.Workbook.Worksheets[i].Cells[k, 2].Value != null)
                    {
                        PriceFrom.Add(Targetpackage.Workbook.Worksheets[i].Cells[k, 2].Value.ToString());
                    }
                }
            }

            
            Targetpackage.Workbook.Worksheets.Add("LastVariant");
            
            
            for (int i = 0; i < ValueFrom.Count-1; i++)
            {
                Targetpackage.Workbook.Worksheets["LastVariant"].Cells[i+1, 4].Value = ValueFrom[i]; // почему столбик 4?????
            }

            for(int i = 0; i < PriceFrom.Count - 1; i++)
            {
                Targetpackage.Workbook.Worksheets["LastVariant"].Cells[i + 1, 6].Value = PriceFrom[i];
            }


            for (int i = 0; i < BooksWithoutNs.Count; i++)
            {
                Targetpackage.Workbook.Worksheets.Delete(BooksWithoutNs[i]);
            }
            
            Targetpackage.Save();
            Console.WriteLine("Work complete");
            Console.ReadLine();
        }
    }
}
