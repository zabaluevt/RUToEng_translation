using ProtocolFieldDefinitionsEditor;
using ProtocolFieldDefinitionsEditor.Classes;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RUToEng_translation
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите 1 для получения всех русских слов");
            Console.WriteLine("Введите 2 перевода слов в документах");
            Console.WriteLine("Введите 3 выделение строки");
            Console.WriteLine("Введите 4 в новый ресурсный файл");
            Console.WriteLine("Введите 5 для получения русских слов из xml файла");
            var parameter = Console.ReadLine();
            switch (parameter)
            {
                case "1":
                    RuToExcel();
                    break;
                case "2":
                    TranslationFiles();
                    break;
                case "3":
                    GetStringInFile();
                    break;
                case "4":
                    AddResourceFile();
                    break;
                case "5":
                    ParseXMLFile();
                    break;
                default:
                    Console.WriteLine("Неправильный ввод");
                    break;
            }
            Console.WriteLine("Программа закончила свою работу!");
            Console.ReadKey();
        }

        private static void ParseXMLFile()
        {
            Console.WriteLine("Введите путь xml файла:");
            string xmlFilePath = Console.ReadLine();

            //string xmlFilePath = @"C:\Users\workstation1\Desktop\sors\UZI\ее\1.xml";

            string str;
            using (StreamReader sr = new StreamReader(File.Open(xmlFilePath, FileMode.Open)))
            {
                str = sr.ReadToEnd();
            }

            // ! Обязательно удалить строчку в xml файле:  xmlns="http://tempuri.org/FormalDocumentFieldDefinitions"
            //

            FormalDocumentFieldDefinitionsCollection collection = EntityDataHelper.Deserialize<FormalDocumentFieldDefinitionsCollection>(str);
            List<string> list = new List<string>();
            Regex regex = new Regex(@"[А-я]");

            foreach (var field in collection.Fields)
            {

                if (regex.IsMatch(field.DisplayedName))
                {
                    CheckIfNotExist(list, field.DisplayedName);
                }

                foreach (var value in field.FormalDocumentFieldValueDefinitions)
                {
                    if (regex.IsMatch(value.DisplayedValue))
                    {
                        CheckIfNotExist(list, value.DisplayedValue);
                    }
                }
            }

            foreach (var field in collection.Groups)
            {
                bool isMatch = regex.IsMatch(field.DisplayedName);
                if (isMatch)
                {
                    CheckIfNotExist(list, field.DisplayedName);
                }
            }

            SaveToExcel(list);
        }

        private static void CheckIfNotExist(List<string> list, string displayedName)
        {
            if (!list.Contains(displayedName))
            {
                list.Add(displayedName);
                System.Console.WriteLine(displayedName);
            }
        }

        private static void AddResourceFile()
        {
            var list = GetExcelStrings();

            RuStrings ruStrings = new RuStrings();

            foreach (var el in list)
            {
                string st = ruStrings.ReadFile(el.Path).Replace(el.Value, "LocalizedStrings.String");
            }
        }

        private static List<PathAndValue> GetExcelStrings()
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(@"C:\Users\workstation1\Desktop\Client.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.UsedRange;

            List<PathAndValue> newList = new List<PathAndValue>();
            for (int index = 1; index < range.Rows.Count; index++)
            {
                string path = ObjWorkSheet.Cells[index, 3].Value.ToString();
                string rustr = ObjWorkSheet.Cells[index, 1].Value.ToString();
                newList.Add(new PathAndValue(path, rustr));
            }
            return newList;
        }

        private static void GetStringInFile()
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open("C:\\Users\\workstation1\\Desktop\\Excel\\TestServices.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.UsedRange;
            for (int index = 1; index < range.Rows.Count; index++)
            {
                string path = ObjWorkSheet.Cells[index, 2].Value.ToString();

                StreamReader sr = new StreamReader(path);

                string[] mass = File.ReadAllLines(path, System.Text.Encoding.Default);

                foreach (var obj in mass)
                {
                    if (obj.Contains(ObjWorkSheet.Cells[index, 1].Value.ToString()))
                    {
                        ObjWorkSheet.Cells[index, 3] = obj;
                        Console.WriteLine("Строка " + index + " СТРОКА: " + obj);
                    }
                        
                }
            }
            ObjWorkBook.SaveAs("C:\\Users\\workstation1\\Desktop\\Excel\\TestServiceWith3str.xlsx");
            ObjExcel.Quit();
        }

        private static void TranslationFiles()
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open("C:\\Users\\workstation1\\Desktop\\Excel\\Test4.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.UsedRange;
            for (int index = 1; index < range.Rows.Count; index++)
            {
                string path =ObjWorkSheet.Cells[index, 2].Value.ToString();

                //StreamReader sr = new StreamReader(path);
                //string file = sr.ReadToEnd();
                //file.Replace(ObjWorkSheet.Cells[index, 1].Value.ToString(), ObjWorkSheet.Cells[index, 3].Value.ToString());
            }
            
            ObjExcel.Quit();

        }

        private static void RuToExcel()
        {
            Console.WriteLine("Введите путь папки:");
            string folderPath = Console.ReadLine();

            DateTime startTime = DateTime.Now;

            RuStrings ruStrings = new RuStrings();

            List<string> listPaths = ruStrings.GetAllFilesInFolder(folderPath);

            foreach (var path in listPaths)
            {
                ruStrings.GetRegexedStrings(path);
            }

            var collection = ruStrings.collection;

            foreach (var item in collection)
                Console.WriteLine(item.Value);

            SaveToExcel(ruStrings.collection);
            Console.WriteLine("Файл сохранен");

            DateTime stopTime = DateTime.Now;

            Console.WriteLine("\nВремя работы программы: " + (stopTime.Second - startTime.Second) + " секунд");
            Console.WriteLine("\nКоличество элементов: " + collection.Count);
        }

        public static void SaveToExcel(List<string> list)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            for (int i = 1; i < list.Count(); i++)
            {
                ObjWorkSheet.Cells[i, 1] = list[i];
            }

            ObjWorkBook.SaveAs("C:\\Users\\workstation1\\Desktop\\Excel\\XML.xlsx");
            ObjExcel.Quit();
        }

        public static void SaveToExcel(Collection<PathAndValue> collection)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            for(int i = 1; i < collection.Count(); i++)
            {
                ObjWorkSheet.Cells[i, 1] = collection[i].Value;
                ObjWorkSheet.Cells[i, 2] = collection[i].Path;
            }

            ObjWorkBook.SaveAs("C:\\Users\\workstation1\\Desktop\\Excel\\TestClient3.xlsx");
            ObjExcel.Quit();
        }
    }
}
