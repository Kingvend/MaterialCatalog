using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace CollisionFinder
{
    class InputOutput
    {
        /// <summary>
        /// Чтение основного файла 
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="LastRow"></param>
        /// <returns></returns>
        public static List<Material> ReaderLargeMaterial(string FilePath, int LastRow, int CodeCol, int NameCol, int FullNameCol_1, int FullNameCol_2, int MessureCol, int CountMesCol)
        {
            List<Material> materialList = new List<Material>();

            Console.WriteLine("Чтение документа");

            using (var fs = new FileStream(FilePath, FileMode.Open))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage doc = new ExcelPackage(fs))
                {
                    Console.WriteLine("Чтение таблицы");

                    ExcelWorksheet sheet = doc.Workbook.Worksheets["Sheet1"];

                    List<Material> tmp_MaterialList = new List<Material>();
                    for (int i = 2; i <= LastRow; i++)
                    {
                        if (sheet.Cells[i, 14] != null)
                        {
                            Material tmp = new Material
                            {
                                MaterialCode = sheet.Cells[i, CodeCol].Value.ToString(),
                                MaterialName = sheet.Cells[i, NameCol].Value.ToString(),
                                MaterialFullName = sheet.Cells[i, FullNameCol_1].Value.ToString() + sheet.Cells[i, FullNameCol_2].Value.ToString(),
                                MaterialMeasureUnit = sheet.Cells[i, MessureCol].Value.ToString(),
                                MaterialCountMU = sheet.Cells[i, CountMesCol].Value.ToString(),
                                MaterialRowNumber = i.ToString(),
                                MaterialSource = Functions.FirstNameFile(FilePath)
                            };

                            tmp_MaterialList.Add(tmp);
                        }
                    }
                    materialList.AddRange(tmp_MaterialList);
                    doc.Dispose();
                }
            }

            return materialList;
        }

        /// <summary>
        /// Чтение комактного файла
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="LastRow"></param>
        /// <returns></returns>
        public static List<Material> ReaderCompactMaterial(string FilePath, int LastRow)
        {
            List<Material> MaterialList = new List<Material>();

            Console.WriteLine("Чтение документа");

            using (var fs = new FileStream(FilePath, FileMode.Open))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage doc = new ExcelPackage(fs))
                {
                    Console.WriteLine("Чтение таблицы");

                    ExcelWorksheet sheet = doc.Workbook.Worksheets["Sheet1"];

                    List<Material> tmp_MaterialList = new List<Material>();
                    for (int i = 1; i < LastRow; i++)
                    {
                        if (sheet.Cells[i, 1] != null)
                        {
                            Material tmp = new Material
                            {
                                MaterialCode = sheet.Cells[i, 1].Value.ToString(),
                                MaterialName = sheet.Cells[i, 2].Value.ToString(),
                                MaterialFullName = sheet.Cells[i, 3].Value.ToString(),
                                MaterialMeasureUnit = sheet.Cells[i, 4].Value.ToString(),
                                MaterialCountMU = sheet.Cells[i, 5].Value.ToString(),
                                MaterialRowNumber = sheet.Cells[i, 6].Value.ToString(),
                                MaterialSource = sheet.Cells[i, 7].Value.ToString()
                            };

                            tmp_MaterialList.Add(tmp);
                        }
                    }
                    MaterialList.AddRange(tmp_MaterialList);
                    doc.Dispose();
                }
            }

            return MaterialList;
        }
        /// <summary>
        /// Чтение нескольких файлов
        /// </summary>
        /// <param name="materialFile"></param>
        /// <returns></returns>
        public static List<Material> ReaderAllMaterial(List<MaterialFile> materialFile)
        {
            List<Material> material = new List<Material>();
            foreach (var mf in materialFile)
            {
                Console.WriteLine("Чтение документа: " + Functions.FirstNameFile(mf.FilePath));
                using (var fs = new FileStream(mf.FilePath, FileMode.Open))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage doc = new ExcelPackage(fs))
                    {
                        Console.WriteLine("Чтение таблицы");

                        ExcelWorksheet sheet = doc.Workbook.Worksheets["Sheet1"];

                        List<Material> tmp_MaterialList = new List<Material>();
                        for (int i = 2; i <= mf.LastRow; i++)
                        {
                            if (sheet.Cells[i, 14] != null)
                            {
                                Material tmp = new Material
                                {
                                    MaterialCode = sheet.Cells[i, mf.CodeCol].Value.ToString(),
                                    MaterialName = sheet.Cells[i, mf.NameCol].Value.ToString(),
                                    MaterialFullName = sheet.Cells[i, mf.FullNameCol_1].Value.ToString() + sheet.Cells[i, mf.FullNameCol_2].Value.ToString(),
                                    MaterialMeasureUnit = sheet.Cells[i, mf.MessureCol].Value.ToString(),
                                    MaterialCountMU = sheet.Cells[i, mf.CountMesCol].Value.ToString(),
                                    MaterialRowNumber = i.ToString(),
                                    MaterialSource = Functions.FirstNameFile(mf.FilePath)
                                };

                                tmp_MaterialList.Add(tmp);
                            }
                        }
                        material.AddRange(tmp_MaterialList);
                        doc.Dispose();
                    }
                }
            }
            return material;
        }
        /// <summary>
        /// Запись компактного файла
        /// </summary>
        /// <param name="material"></param>
        /// <param name="NewFilePath"></param>
        public static void WriterMaterial(List<Material> material, string NewFilePath)
        {
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet = doc.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].LoadFromCollection(material);
                FileInfo fi = new FileInfo(NewFilePath);
                doc.SaveAs(fi);
                doc.Dispose();
            }
        }
        /// <summary>
        /// создание excel файла с коллизиями по коду
        /// </summary>
        /// <param name="collision"></param>
        /// <param name="fileName"></param>

        public static void WriterCodeCollision(List<Collision> collision, string fileName, int firstRow)
        {
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet = doc.Workbook.Worksheets.Add("Sheet1");
                int numberRow = firstRow;
                var GroupCodeMaterial = collision
                    .GroupBy(s => s.Code);
                foreach (var s1 in GroupCodeMaterial)
                {
                    sheet.Cells[numberRow, 1].Value = s1.Key.ToString();
                    foreach (var s2 in s1)
                    {
                        sheet.Cells[numberRow, 2].Value = s2.Name.ToString();
                        sheet.Cells[numberRow, 3].Value = s2.FullName.ToString();
                        sheet.Cells[numberRow, 4].Value = s2.FileSource;
                        sheet.Cells[numberRow, 2].Value = s2.RowNumber.ToString();
                        numberRow++;
                    }
                }
                FileInfo fi = new FileInfo(fileName);
                doc.SaveAs(fi);
                doc.Dispose();
            }
        }
        public static void WriterNameCollision(List<Collision> collision, string fileName, int firstRow)
        {
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet = doc.Workbook.Worksheets.Add("Sheet1");
                int numberRow = firstRow;
                var GroupNameMaterial = collision
                    .GroupBy(s => s.Name);
                foreach (var s1 in GroupNameMaterial)
                {
                    sheet.Cells[numberRow, 1].Value = s1.Key.ToString();
                    foreach (var s2 in s1)
                    {
                        sheet.Cells[numberRow, 2].Value = s2.Code.ToString();
                        sheet.Cells[numberRow, 3].Value = s2.FullName.ToString();
                        sheet.Cells[numberRow, 4].Value = s2.FileSource;
                        sheet.Cells[numberRow, 5].Value = s2.RowNumber.ToString();
                        numberRow++;
                    }
                }
                FileInfo fi = new FileInfo(fileName);
                doc.SaveAs(fi);
                doc.Dispose();
            }
        }

        public static List<MTR_Catalog> ReadMaterialForCatalog(List<MtrCatalogFileProperty> mtrCatalog)
        {
            var mtrList = new List<MTR_Catalog>();

            foreach (var ml2 in mtrCatalog)
            {
                var ml = ml2;
                Console.WriteLine("Чтение документа: " + Functions.FirstNameFile(ml.FilePath));
                using (var fs = new FileStream(ml.FilePath, FileMode.Open))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage doc = new ExcelPackage(fs))
                    {
                        Console.WriteLine("Чтение таблицы");

                        ExcelWorksheet sheet = doc.Workbook.Worksheets["Sheet1"];
                        ml = ml.FindColumns(ref sheet, ml);
                        var tmp_MTRCatalogList = new List<MTR_Catalog>();
                        //var prop = typeof(MtrCatalogFileProperty).GetProperties();
                        //for (int i = 0; i < prop.Length; i++)
                        //{
                        //    Console.WriteLine(prop[i]. .ToString());

                        //}
                        for (int i = ml.FirstRow; i <= ml.LastRow; i++)
                        {
                            var tmp = new MTR_Catalog
                            {
                                MaterialCode = sheet.Cells[i, ml.MaterialCodeCol].Value.ToString(),

                                BlockCode = Functions.BlockCodeConvert(sheet.Cells[i, ml.BlockCodeCol].Value.ToString()),

                                MaterialName = sheet.Cells[i, ml.MaterialNameCol].Value.ToString(),

                                MaterialFullName = sheet.Cells[i, ml.MaterialFullName1Col].Value.ToString() + sheet.Cells[i, ml.MaterialFullName2Col].Value.ToString(),

                                GroupName = sheet.Cells[i, ml.GroupNameCol].Value.ToString(),

                                GroupCode = sheet.Cells[i, ml.GroupCodeCol].Value.ToString(),

                                NaimCodeClass = sheet.Cells[i, ml.NaimCodeClassCol].Value.ToString(),

                                ConsigneeDetail = sheet.Cells[i, ml.ConsigneeDetailCol].Value.ToString(),

                                DeliveryDate = sheet.Cells[i, ml.DeliveryDateCol].Value.ToString(),

                                BasisMU = sheet.Cells[i, ml.BasisMUCol].Value.ToString(),

                                BasisMUCount = sheet.Cells[i, ml.BasisMUCountCol].Value.ToString(),

                                BasisMUPrice = sheet.Cells[i, ml.BasisMUPriceCol].Value.ToString(),

                                AltMU = sheet.Cells[i, ml.AltMUCol].Value.ToString(),

                                AltMUCount = sheet.Cells[i, ml.AltMUCountCol].Value.ToString(),

                                AltMUPrice = sheet.Cells[i, ml.AltMUPriceCol].Value.ToString(),

                                SPPName = sheet.Cells[i, ml.SPPNameCol].Value.ToString(),

                                SPPElem = sheet.Cells[i, ml.SPPElemCol].Value.ToString(),

                                OKPD2 = sheet.Cells[i, ml.OKPD2Col].Value.ToString(),

                                OKPD2Code = sheet.Cells[i, ml.OKPD2CodeCol].Value.ToString(),

                                Brutto = sheet.Cells[i, ml.BruttoCol].Value.ToString(),

                                Kol_voSCHF = sheet.Cells[i, ml.Kol_voSCHFCol].Value.ToString(),

                                SumSCHFWithoutNDS = sheet.Cells[i, ml.SumSCHFWithoutNDSCol].Value.ToString(),

                                DateSchf = (sheet.Cells[i, ml.DateSchfCol].Value ?? string.Empty).ToString()
                            };
                            tmp_MTRCatalogList.Add(tmp);

                        }
                        mtrList.AddRange(tmp_MTRCatalogList);
                        doc.Dispose();
                    }
                }
            }
            return mtrList;
        }

  
    }

    
}
