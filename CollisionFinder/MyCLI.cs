using System;
using System.Collections.Generic;
using System.Linq;

namespace CollisionFinder
{
    class MyCLI
    {
        public static void Menu()
        {
            bool isExit = false;
            int countNotesInMateriaList = 0;
            while (!isExit)
            {
                string com = "";
                Console.Clear();
                Console.WriteLine("Обработка коллизий");
                Console.WriteLine("-------------------------------------------");
                Console.WriteLine();
                Console.WriteLine("Доступные функции");
                Console.WriteLine();
                Console.WriteLine("1) загрузка новых excel файлов");
                Console.WriteLine("2) Сформировать отчет о коллизиях");
                Console.WriteLine("0) Выход из программы");
                Console.WriteLine("-------------------------------------------");
                Console.WriteLine("Выбирете доступную команду");
                com = Console.ReadLine();
                if (com == "1")
                {
                    countNotesInMateriaList = Com1();
                }
                else
                {
                    if (com == "2")
                    {
                        if(countNotesInMateriaList != 0)
                        {
                            Com2(countNotesInMateriaList);
                        }
                        else
                        {
                            Console.WriteLine("Ошибка");
                            Console.ReadKey();
                        }
                        
                    }
                    else
                    {
                        if (com == "0")
                        {
                            isExit = Com0();
                        }
                        else
                        {
                            Console.WriteLine("Ошибка! Не существующая команда.");
                        }
                    }
                }
            }

        }

        static int Com1()
        {
            int countNotes = 0;
            List<MaterialFile> materialFiles = new List<MaterialFile>();
            List<Material> materialList = new List<Material>();
            string newPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\TotalFile.xlsx";
            bool isAllFiles = false;
            while (isAllFiles == false)
            {
                bool isDataNorm = false;
                string path = "";
                string lastRow = "";
                string codeCol = "";
                string nameCol = "";
                string fullName1Col = "";
                string fullName2Col = "";
                string messureCol = "";

                while (isDataNorm == false)
                {
                    path = "";
                    lastRow = "";
                    codeCol = "";
                    nameCol = "";
                    fullName1Col = "";
                    fullName2Col = "";
                    messureCol = "";
                    Console.Clear();
                    Console.WriteLine("Введите полный путь добавляемого файла:");
                    path = Console.ReadLine();
                    Console.WriteLine("Введите номер последней колонки:");
                    lastRow = Console.ReadLine();
                    Console.WriteLine("Введите имя столбца, содержащий код материала:");
                    codeCol = Console.ReadLine();
                    Console.WriteLine("Введите имя столбца, содержащий краткое наименование материала:");
                    nameCol = Console.ReadLine();
                    Console.WriteLine("Введите имя столбца, содержащий полное наименование материала(1):");
                    fullName1Col = Console.ReadLine();
                    Console.WriteLine("Введите имя столбца, содержащий полное наименование материала(2):");
                    fullName2Col = Console.ReadLine();
                    Console.WriteLine("Введите имя столбца, содержащий единицы измерения материала:");
                    messureCol = Console.ReadLine();
                    Console.Clear();
                    Console.WriteLine("Введены следующие данные: ");
                    Console.WriteLine("Полный путь добавляемого файла: " + path);
                    Console.WriteLine("Номер последней колонки: " + lastRow);
                    Console.WriteLine("Имя столбца, содержащий код материала: " + codeCol);
                    Console.WriteLine("Имя столбца, содержащий краткое наименование материала: " + nameCol);
                    Console.WriteLine("Имя столбца, содержащий полное наименование материала(1): " + fullName1Col);
                    Console.WriteLine("Имя столбца, содержащий полное наименование материала(2): " + fullName2Col);
                    Console.WriteLine("Имя столбца, содержащий единицы измерения материала: " + messureCol);
                    Console.WriteLine();
                    Console.WriteLine("Если данные верны, нажмите Y(Д), если необходимо ввести изменить данные нажмите N(Н)");
                    string ansCom = Console.ReadLine();
                    if ((ansCom == "Y") || (ansCom == "Д"))
                    {
                        isDataNorm = true;
                    }
                    else
                    {
                        if ((ansCom == "N") || (ansCom == "Н"))
                        {
                            isDataNorm = false;
                        }
                    }
                }
                countNotes += Int32.Parse(lastRow) - 1;
                MaterialFile tmp_mf = new MaterialFile
                {
                    FilePath = path,
                    LastRow = Int32.Parse(lastRow),
                    CodeCol = Functions.ConvertNumberColumnInExcel(codeCol),
                    NameCol = Functions.ConvertNumberColumnInExcel(nameCol),
                    FullNameCol_1 = Functions.ConvertNumberColumnInExcel(fullName1Col),
                    FullNameCol_2 = Functions.ConvertNumberColumnInExcel(fullName2Col),
                    MessureCol = Functions.ConvertNumberColumnInExcel(messureCol),
                    CountMesCol = Functions.ConvertNumberColumnInExcel(codeCol) // ITS BAD 
                };
                materialFiles.Add(tmp_mf);
                Console.WriteLine("Файл " + Functions.FirstNameFile(path) + " добавлен в обработку");
                Console.WriteLine("Добавить еще файлы? (Y/N) или (Д/Н)");
                string ans_2 = Console.ReadLine();
                if ((ans_2 == "Y") || (ans_2 == "Д"))
                {
                    isAllFiles = false;
                }
                else
                {
                    if ((ans_2 == "N") || (ans_2 == "Н"))
                    {
                        isAllFiles = true;
                    }
                }
            }
            materialList = InputOutput.ReaderAllMaterial(materialFiles);
            Console.WriteLine("Чтение файла завершено");
            InputOutput.WriterMaterial(materialList, newPath_3); // Запись в компактный excel файл
            Console.ReadKey();
            return countNotes;
        }
        static void Com2(int countNotes)
        {
            List<Material> materialList = new List<Material>();
            List<Collision> codeCollision = new List<Collision>();
            List<Collision> nameCollision = new List<Collision>();
            List<Collision> messureCollision = new List<Collision>();
            List<Collision> fullNameCollision = new List<Collision>();

            string newPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\TotalFile.xlsx";
            materialList = InputOutput.ReaderCompactMaterial(newPath_3, countNotes);

            // проверка по коду материала

            var groupCodeMaterial = materialList
                .GroupBy(s => s.MaterialCode);
            foreach (var s1 in groupCodeMaterial)
            {
                int count = 0;
                List<Collision> _CodeCollision = new List<Collision>();
                foreach (var s2 in s1)
                {
                    if (count == 0)
                    {
                        Collision collision = new Collision
                        {
                            Code = s1.Key,
                            Name = s2.MaterialName,
                            FullName = s2.MaterialFullName,
                            RowNumber = s2.MaterialRowNumber,
                            FileSource = s2.MaterialSource,
                            Messure = s2.MaterialMeasureUnit
                        };
                        _CodeCollision.Add(collision);
                        count++;
                    }
                    else
                    {
                        int countCollision = 0;
                        foreach (var c in _CodeCollision)
                        {
                            if ((Functions.ChangeString(c.Name) != Functions.ChangeString(s2.MaterialName)) || (Functions.ChangeString(c.FullName) != Functions.ChangeString(s2.MaterialFullName)))
                            {
                                countCollision++;
                            }
                            count++;
                        }
                        if (countCollision == _CodeCollision.Count)
                        {
                            Collision collision = new Collision
                            {
                                Code = s1.Key,
                                Name = s2.MaterialName,
                                FullName = s2.MaterialFullName,
                                RowNumber = s2.MaterialRowNumber,
                                FileSource = s2.MaterialSource,
                                Messure = s2.MaterialMeasureUnit
                            };
                            _CodeCollision.Add(collision);
                        }

                    }
                }
                if (_CodeCollision.Count > 1)
                {
                    codeCollision.AddRange(_CodeCollision);
                }
            }

            Console.WriteLine("Проверка по коду материала завершена");

            // проверка по наименованию материала
            var groupNameMaterial = materialList
                .GroupBy(s => s.MaterialName);
            foreach (var s1 in groupNameMaterial)
            {
                int count = 0;
                int difCount = 0;
                List<Collision> _NameCollision = new List<Collision>();

                foreach (var s2 in s1)
                {
                    if (count == 0)
                    {
                        Collision collision = new Collision
                        {
                            Code = s2.MaterialCode,
                            Name = s1.Key,
                            FullName = s2.MaterialFullName,
                            RowNumber = s2.MaterialRowNumber,
                            FileSource = s2.MaterialSource,
                            Messure = s2.MaterialMeasureUnit
                        };
                        _NameCollision.Add(collision);
                        count++;
                        difCount++;
                    }
                    else
                    {
                        int countCollision = 0;
                        for (var c = 0;c < _NameCollision.Count; c++)
                        {
                            if (_NameCollision[c].Code == s2.MaterialCode && _NameCollision[c].Name == s2.MaterialName && _NameCollision[c].FullName == s2.MaterialFullName)
                            {
                                Collision collision = new Collision
                                {
                                    Code = s2.MaterialCode,
                                    Name = s1.Key,
                                    FullName = s2.MaterialFullName,
                                    RowNumber = s2.MaterialRowNumber,
                                    FileSource = s2.MaterialSource
                                };
                                _NameCollision.Add(collision);
                                break;
                            }
                            else
                            {
                                if ((_NameCollision[c].Code != s2.MaterialCode))
                                {
                                    if ((_NameCollision[c].FullName == s2.MaterialFullName))
                                    {
                                        countCollision++;
                                    }
                                }
                            }
                            
                        }
                        if (countCollision == _NameCollision.Count)
                        {
                            Collision collision = new Collision
                            {
                                Code = s2.MaterialCode,
                                Name = s1.Key,
                                FullName = s2.MaterialFullName,
                                RowNumber = s2.MaterialRowNumber,
                                FileSource = s2.MaterialSource
                            };
                            _NameCollision.Add(collision);
                            difCount++;
                        }
                        count++;
                    }
                }
                if (difCount > 1)
                {
                    nameCollision.AddRange(_NameCollision);
                }
            }
            Console.WriteLine("Проверка по наименованию материала завершена");

            // тест новой проверки (конец)

            // проверка единиц измерения

            var groupMessureMaterial = materialList
                .GroupBy(s => s.MaterialName);
            foreach (var s1 in groupMessureMaterial)
            {
                int count = 0;
                int difCount = 0;
                List<Collision> _MessureCollision = new List<Collision>();


                foreach (var s2 in s1)
                {
                    if (count == 0)
                    {
                        Collision collision = new Collision
                        {
                            Code = s2.MaterialCode,
                            Name = s1.Key,
                            FullName = s2.MaterialFullName,
                            RowNumber = s2.MaterialRowNumber,
                            FileSource = s2.MaterialSource,
                            Messure = s2.MaterialMeasureUnit
                        };
                        _MessureCollision.Add(collision);
                        count++;
                        difCount++;
                    }
                    else
                    {
                        int countCollision = 0;                       
                        for(int c = 0; c < _MessureCollision.Count;c++)
                        {
                            if ((_MessureCollision[c].FullName == s2.MaterialFullName) && (Functions.ChangeString(_MessureCollision[c].Messure) == Functions.ChangeString(s2.MaterialMeasureUnit)))
                                {
                                Collision collision = new Collision
                                {
                                    Code = s2.MaterialCode,
                                    Name = s1.Key,
                                    FullName = s2.MaterialFullName,
                                    RowNumber = s2.MaterialRowNumber,
                                    FileSource = s2.MaterialSource,
                                    Messure = s2.MaterialMeasureUnit
                                };
                                _MessureCollision.Add(collision);
                                break;
                            }
                            else
                            {
                                if ((_MessureCollision[c].FullName == s2.MaterialFullName) && (Functions.ChangeString(_MessureCollision[c].Messure) != Functions.ChangeString(s2.MaterialMeasureUnit)))
                                {
                                    countCollision++;
                                }
                            }
                        }
                        if (countCollision == _MessureCollision.Count)
                        {
                            Collision collision = new Collision
                            {
                                Code = s2.MaterialCode,
                                Name = s1.Key,
                                FullName = s2.MaterialFullName,
                                RowNumber = s2.MaterialRowNumber,
                                FileSource = s2.MaterialSource,
                                Messure = s2.MaterialMeasureUnit
                            };
                            _MessureCollision.Add(collision);
                            difCount++;
                        }
                        count++;
                    }
                }
                if (difCount > 1)
                {
                    messureCollision.AddRange(_MessureCollision);
                }
            }
            Console.WriteLine("Проверка по наименованию единиц измерения материала завершена");

            // проверка по полному наименованию материала
            var groupFullNameMaterial = materialList
                .GroupBy(s => s.MaterialFullName);
            foreach (var s1 in groupFullNameMaterial)
            {
                int count = 0;
                int difCount = 0;
                List<Collision> _FullNameCollision = new List<Collision>();


                foreach (var s2 in s1)
                {
                    if (count == 0)
                    {
                        Collision collision = new Collision
                        {
                            Code = s2.MaterialCode,
                            Name = s2.MaterialName,
                            FullName = s1.Key,
                            RowNumber = s2.MaterialRowNumber,
                            FileSource = s2.MaterialSource,
                            Messure = s2.MaterialMeasureUnit
                        };
                        _FullNameCollision.Add(collision);
                        count++;
                        difCount++;
                    }
                    else
                    {
                        int countCollision = 0;
                        for (var c = 0; c < _FullNameCollision.Count; c++)
                        {
                            if (_FullNameCollision[c].Code == s2.MaterialCode && _FullNameCollision[c].Name == s2.MaterialName && _FullNameCollision[c].FullName == s2.MaterialFullName)
                            {
                                Collision collision = new Collision
                                {
                                    Code = s2.MaterialCode,
                                    Name = s2.MaterialName,
                                    FullName = s1.Key,
                                    RowNumber = s2.MaterialRowNumber,
                                    FileSource = s2.MaterialSource
                                };
                                _FullNameCollision.Add(collision);
                                break;
                            }
                            else
                            {
                                if ((_FullNameCollision[c].Code == s2.MaterialCode))
                                {
                                    if ((_FullNameCollision[c].Name != s2.MaterialName))
                                    {
                                        countCollision++;
                                    }
                                }
                            }

                        }
                        if (countCollision == _FullNameCollision.Count)
                        {
                            Collision collision = new Collision
                            {
                                Code = s2.MaterialCode,
                                Name = s2.MaterialName,
                                FullName = s1.Key,
                                RowNumber = s2.MaterialRowNumber,
                                FileSource = s2.MaterialSource
                            };
                            _FullNameCollision.Add(collision);
                            difCount++;
                        }
                        count++;
                    }
                }
                if (difCount > 1)
                {
                    nameCollision.AddRange(_FullNameCollision);
                }
            }
            Console.WriteLine("Проверка по наименованию материала завершена");

            // тест новой проверки (конец)

            string nameColPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\Name_Collision";
            string codeColPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\Code_Collision";
            string reportPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\report_3";

            InputOutput.WriterCodeCollision(codeCollision, codeColPath_3 + "_2.xlsx", 4);
            InputOutput.WriterNameCollision(nameCollision, nameColPath_3 + "_2.xlsx", 4);
            Report.ReportGenerate(codeCollision, nameCollision, messureCollision, fullNameCollision, 1, reportPath_3 + ".xlsx");

            Console.WriteLine("ГОТОВО!");
            Console.ReadKey();
        }
        static bool Com0()
        {
            return true;
        }
    }
}



