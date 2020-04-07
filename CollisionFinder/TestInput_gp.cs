//using GazpromNeft.Excel.Document.Common;
//using GazpromNeft.Excel.Document.EPPlusImpl;
using OfficeOpenXml;

using System;
using System.Collections.Generic;

namespace CollisionFinder
{
    class TestInput_gp
    {       
        public static void Reader()
        {
            IExcelDocument doc = EPExcelDocument.FromFile(@"D:\test.xlsx");

            IExcelSheet sheet = doc.SheetsByName["sheet1"];

            ITableReader table = sheet.GetTableReader("A", "Z", 1, 10);


            object hight = 10m;


            table.ReadLines(row =>
            {
                Console.WriteLine($"value=\t{row[1]}");
            });

            doc.Dispose();
        }

        public static List<Material> ReaderLargeMaterial(string FilePath, int LastRow)
        {
            List<Material> MaterialList = new List<Material>();

            Console.WriteLine("Чтение документа");
            IExcelDocument doc = EPExcelDocument.FromFile(FilePath);

            Console.WriteLine("Чтение таблицы");
            IExcelSheet sheet = doc.SheetsByName["Sheet1"];

            ITableReader table = sheet.GetTableReader("N", "O", 2, LastRow);
            ITableReader table_2 = sheet.GetTableReader("GL", "GL", 2, LastRow);
            ITableReader table_3 = sheet.GetTableReader("GO", "GO", 2, LastRow);

            List<Material> tmp_MaterialList = new List<Material>();


            table.ReadLines(row =>
            {
                Material _tmp_material = new Material
                {
                    MaterialCode = (string)row[1],
                    MaterialName = (string)row[2]
            };
                tmp_MaterialList.Add(_tmp_material);
            });

            List<string> fullMaterialName_1 = new List<string>();
            table_2.ReadLines(row =>
            {              
                {
                    fullMaterialName_1.Add((string)row[1]);
                };
            });

            List<string> fullMaterialName_2 = new List<string>();
            table_3.ReadLines(row =>
            {
                {
                    fullMaterialName_2.Add((string)row[1]);
                };
            });
            //сборка времменых списков материалов
            MaterialList.AddRange(tmp_MaterialList);
            for(int i = 0; i < MaterialList.Count; i++)
            {              
                MaterialList[i].MaterialFullName = fullMaterialName_1[i] + fullMaterialName_2[i];
            }

            Console.WriteLine("Чтение завершено");
            doc.Dispose();
            return MaterialList;
        }
        public static List<Material> ReaderCompactMaterial(string FilePath, int LastRow)
        {
            List<Material> MaterialList = new List<Material>();

            Console.WriteLine("Чтение документа");
            IExcelDocument doc = EPExcelDocument.FromFile(FilePath);

            Console.WriteLine("Чтение таблицы");
            IExcelSheet sheet = doc.SheetsByName["sheet1"];

            ITableReader table = sheet.GetTableReader("A", "C", 1, LastRow);

            table.ReadLines(row =>
            {
                Material _tmp_material = new Material
                {
                    MaterialCode = (string)row[1],
                    MaterialName = (string)row[2],
                    MaterialFullName = (string)row[3]
                };
                MaterialList.Add(_tmp_material);
            });
            Console.WriteLine("Чтение завершено");
            doc.Dispose();
            return MaterialList;
        }

        public static void WriterMaterial(List<Material> material, string NewFilePath)
        {
            IExcelDocument doc = EPExcelDocument.NewFile();

            IExcelSheet sheet = doc.Sheets[1];

            ITableWriter table = sheet.GetTableWriter("A", "С", 1, new TableWriterSettings());

            for (int i = 1; i <= material.Count; i++)
            {
                table.AddRow(row =>
                {
                    row[1].Value = material[i-1].MaterialCode;
                    row[2].Value = material[i-1].MaterialName;
                    row[3].Value = material[i - 1].MaterialFullName;
                });
            }
            doc.SaveAs(NewFilePath);

            doc.Dispose();
        }
        public static void Writer()
        {
            IExcelDocument doc = EPExcelDocument.NewFile();

            IExcelSheet sheet = doc.Sheets[1];

            ITableWriter table = sheet.GetTableWriter("A", "Z", 1, new TableWriterSettings());

            for (int i = 1; i < 11; i++)
            {
                table.AddRow(row =>
                {
                    row[1].Value = $"test{i.ToString("D2")}";
                    row[2].Value = $"cell_{i}";

                    row[1].Bold = true;

                    row.Italic = true;

                    row[2].MergeColumnSpan = 2;
                });
            }

            doc.SaveAs(@"D:\test.xlsx");

            doc.Dispose();

        }
    }
}
