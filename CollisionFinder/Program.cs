using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace CollisionFinder
{

    class Program
    {
        static void Main(string[] args)
        {
            //string path = @"C:\Users\Alex\Desktop\файлы\ДВ ГПН Ямал 16.03.2020.zip\ДВ ГПН Ямал 16.03.2020_test.xlsx";
            //string newPath = @"C:\Users\Alex\Desktop\файлы\ДВ ГПН Ямал 16.03.2020.zip\ДВ ГПН Ямал 16.03.2020_test_2.xlsx";

            //string path_2 = @"C:\Users\Alex\Desktop\файлы\Выгрузка 509 17.03.2020.zip\Выгрузка 509 17.03.2020_test.xlsx";
            //string newPath_2 = @"C:\Users\Alex\Desktop\файлы\Выгрузка 509 17.03.2020.zip\Выгрузка 509 17.03.2020_test_2.xlsx";

            //string newPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\TotalFile.xlsx";
            //  MyCLI.Menu(); 
            Test();


        }
        static void Test()
        {
            string fileName = @"C:\Users\Alex\Desktop\ttt.xlsx";
            var mtrProp = new List<MtrCatalogFileProperty>();

            var MtrCatalogFileProperty = new MtrCatalogFileProperty
            {
                FilePath = @"C:\Users\Alex\Desktop\файлы\ДВ ГПН Ямал 16.03.2020.zip\ДВ ГПН Ямал 16.03.2020_test.xlsx",

                FirstRow = 2,

                LastRow = 33999,
            };
            mtrProp.Add(MtrCatalogFileProperty);

            MtrCatalogFileProperty = new MtrCatalogFileProperty
            {
                FilePath = @"C:\Users\Alex\Desktop\файлы\Выгрузка 509 17.03.2020.zip\Выгрузка 509 17.03.2020_test.xlsx",

                FirstRow = 2,

                LastRow = 106772,
            };
            mtrProp.Add(MtrCatalogFileProperty);

            var MaterialCatalog = InputOutput.ReadMaterialForCatalog(mtrProp);

            double minSize = 75;
            double maxSize = 150;
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet = doc.Workbook.Worksheets.Add("Sheet1");
                MTR_Catalog.Header(sheet, MaterialCatalog);

                for(int i = 1; i <= 18; i++)
                {
                    if (i == 4)
                    {
                        sheet.Column(i).AutoFit(minSize, maxSize);
                        //continue;
                    }
                    else
                    {
                        if (i == 5)
                        {
                            continue;
                        }
                        else
                        {
                            sheet.Column(i).AutoFit();                            
                        }
                    }
                    sheet.Column(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }
               
                FileInfo fi = new FileInfo(fileName);

                doc.SaveAs(fi);
                doc.Dispose();
            }
        }

        ////static void TestProgram() // тестирование программы
        ////{
        ////    List<Group> MainList = new List<Group>(); // список всех групп

        ////    var grouped = materialList.GroupBy(s => PrepareString(s));


        ////    foreach (var grp in grouped)
        ////    {
        ////        var item = grp;
        ////    }

        ////    bool flag = true;
        ////    string b;
        ////    while (flag)
        ////    {
        ////        b = Console.ReadLine(); // получение новой строки
        ////        if (b.Length == 0)
        ////        {
        ////            flag = false;
        ////        }
        ////        if (b == "1")
        ////        {
        ////            Output(MainList);
        ////        }
        ////        else
        ////        {
        ////            if (b == "2")
        ////            {
        ////                foreach (var l in materialList)
        ////                {
        ////                    if (MainList.Count == 0)
        ////                    {
        ////                        AddNewNote(l, ref MainList);
        ////                    }
        ////                    else
        ////                    {
        ////                        if (IsStringExiting(l, ref MainList))
        ////                        {
        ////                            Console.WriteLine("Строка существует! число увеличено");
        ////                        }
        ////                        else
        ////                        {
        ////                            CompaireByGroups(4, l, MainList);
        ////                            //сравнение с группами

        ////                        }
        ////                    }
        ////                }
        ////            }
        ////            else
        ////            {
        ////                if (MainList.Count == 0)
        ////                {
        ////                    AddNewNote(b, ref MainList);
        ////                }
        ////                else
        ////                {
        ////                    if (IsStringExiting(b, ref MainList))
        ////                    {
        ////                        Console.WriteLine("Строка существует! число увеличено");
        ////                    }
        ////                    else
        ////                    {
        ////                        CompaireByGroups(4, b, MainList);
        ////                    }
        ////                }
        ////            }
        ////        }
        ////    }
        ////}
        //private static string PrepareString(string s)
        //{
        //    var result = s.ToLower().Split(' ');
        //    return result[0];
        //}
        ///// <summary>
        ///// тестовая функция вывода группировки по наименованию материала 
        ///// </summary>
        ///// <param name="MainList"></param>
        //static void Output(List<Group> MainList)
        //{
        //    Console.WriteLine("--------------------------------------------");
        //    Console.WriteLine("Список групп");
        //    for (int i = 0; i < MainList.Count; i++)
        //    {
        //        Console.WriteLine("--------------------------------------------");
        //        Console.Write("Группа №{0} ", i + 1);
        //        Console.WriteLine("число различных вариаций: {0}", MainList[i].StringInGroup.Count);
        //        Console.WriteLine("Имя: {0}", MainList[i].AverageNote.OriginalString);
        //        Console.WriteLine();
        //        for (int j = 0; j < MainList[i].StringInGroup.Count; j++)
        //        {
        //            Console.WriteLine("\t" + MainList[i].StringInGroup[j].OriginalString);
        //            Console.WriteLine("\t\tЧисло повторений:" + MainList[i].StringInGroup[j].CountStringInNote);
        //        }
        //    }
        //    Console.WriteLine("--------------------------------------------");
        //}
        //static void AddNewNote(string NewString, ref List<Group> MainList) // добавление новой записи в главный список + (пополнение числа однотипных)
        //{
        //    Group tmp_group = new Group
        //    {
        //        StringInGroup = new List<Note>()
        //    };
        //    Note tmp_note = new Note
        //    {
        //        OriginalString = NewString,

        //        ChangedString = NewChangeString(NewString),
        //        CountStringInNote = 1
        //    };
        //    tmp_group.StringInGroup.Add(tmp_note);
        //    tmp_group.AverageNote = tmp_note;
        //    MainList.Add(tmp_group);
        //}

        //static bool IsStringExiting(string s, ref List<Group> MainList)  // проверка наличия данной строки в списке + increment
        //{
        //    bool status = false;

        //    for (var i = 0; i < MainList.Count; i++)
        //    {
        //        for (var j = 0; j < MainList[i].StringInGroup.Count; j++)
        //        {
        //            if (s == MainList[i].StringInGroup[j].OriginalString)
        //            {
        //                Note tmp = new Note
        //                {
        //                    OriginalString = MainList[i].StringInGroup[j].OriginalString,
        //                    ChangedString = MainList[i].StringInGroup[j].ChangedString,
        //                    CountStringInNote = MainList[i].StringInGroup[j].CountStringInNote + 1
        //                };
        //                MainList[i].StringInGroup[j] = tmp;

        //                Group RefOnGroup = MainList[i];
        //                UpdateAverageString(MainList[i].StringInGroup[j], ref RefOnGroup);

        //                status = true;
        //                return status;
        //            }
        //        }
        //    }
        //    return status;
        //}
        //static Note FindAverageString(ref Group ln) // поиск среднего значения группы
        //{
        //    Note AnsNote = ln.AverageNote;
        //    int Max = -1;

        //    for (int i = 0; i < ln.StringInGroup.Count; i++)
        //    {
        //        if (ln.StringInGroup[i].CountStringInNote > Max)
        //        {
        //            AnsNote = ln.StringInGroup[i];
        //            Max = ln.StringInGroup[i].CountStringInNote;
        //        }
        //    }
        //    return AnsNote;
        //}
        //static void UpdateAverageString(Note n, ref Group g) // обновление усредненной записи
        //{
        //    if (n.CountStringInNote > g.AverageNote.CountStringInNote)
        //    {
        //        g.AverageNote = n;
        //    }
        //}
        //static void CompaireByGroups(int MaxSubStr, string NewString, List<Group> MainList) // сравнение по группам
        //{
        //    double Max = -1;
        //    int iter = 0;
        //    string iter_string = "";
        //    for (int i = 0; i < MainList.Count; i++)
        //    {
        //        double tmp_max = IndistinctMatching(MaxSubStr, NewChangeString(NewString), MainList[i].AverageNote.ChangedString);
        //        if (tmp_max > Max)
        //        {
        //            Max = tmp_max;
        //            iter = i;
        //            iter_string = MainList[i].AverageNote.OriginalString;
        //        }
        //    }
        //    //Console.WriteLine("--------------------------------------------");
        //    //Console.WriteLine("Анализ группы");
        //    //Console.WriteLine("Наиболшее совпадение с новой строкой имеет запись {0}: {1}: {2}% совпадения",iter, iter_string, Max);
        //    //Console.WriteLine("--------------------------------------------");
        //    if (Max < 90)
        //    {
        //        AddNewNote(NewString, ref MainList);
        //    }
        //    else
        //    {
        //        Note tmp = new Note
        //        {
        //            OriginalString = NewString,
        //            ChangedString = NewChangeString(NewString),
        //            CountStringInNote = 1
        //        };
        //        MainList[iter].StringInGroup.Add(tmp);
        //    }
        //}
        //static RetCount Matching(string strInputA, string strInputB, int lngLen)
        //{
        //    RetCount TempRet = new RetCount();
        //    int PosStrA;
        //    int PosStrB;
        //    string strTempA;
        //    string strTempB;
        //    TempRet.lngCountLike = 0;
        //    TempRet.lngSubRows = 0;
        //    for (PosStrA = 0; PosStrA <= strInputA.Length - lngLen; PosStrA++)
        //    {
        //        strTempA = strInputA.Substring(PosStrA, lngLen);
        //        for (PosStrB = 0; PosStrB <= strInputB.Length - lngLen; PosStrB++)
        //        {
        //            strTempB = strInputB.Substring(PosStrB, lngLen);
        //            if ((string.Compare(strTempA, strTempB) == 0))
        //            {
        //                TempRet.lngCountLike = (TempRet.lngCountLike + 1);
        //                break;
        //            }
        //        }
        //        TempRet.lngSubRows = (TempRet.lngSubRows + 1);
        //    }
        //    return TempRet;
        //}
        //static double IndistinctMatching(int MaxMatching, string strInputMatching, string strInputStandart)
        //{
        //    RetCount gret = new RetCount();
        //    RetCount tret = new RetCount();
        //    int lngCurLen; //текущая длина подстроки
        //    //если не передан какой-либо параметр, то выход
        //    if (MaxMatching == 0 || strInputMatching.Length == 0 || strInputStandart.Length == 0) return 0;
        //    gret.lngCountLike = 0;
        //    gret.lngSubRows = 0;
        //    // Цикл прохода по длине сравниваемой фразы
        //    for (lngCurLen = 1; lngCurLen <= MaxMatching; lngCurLen++)
        //    {
        //        //Сравниваем строку A со строкой B
        //        tret = Matching(strInputMatching, strInputStandart, lngCurLen);
        //        gret.lngCountLike = gret.lngCountLike + tret.lngCountLike;
        //        gret.lngSubRows = gret.lngSubRows + tret.lngSubRows;
        //        //Сравниваем строку B со строкой A
        //        tret = Matching(strInputStandart, strInputMatching, lngCurLen);
        //        gret.lngCountLike = gret.lngCountLike + tret.lngCountLike;
        //        gret.lngSubRows = gret.lngSubRows + tret.lngSubRows;
        //    }
        //    if (gret.lngSubRows == 0) return 0;
        //    return (double)(gret.lngCountLike * 100.0 / gret.lngSubRows);
        //}
    }
}



