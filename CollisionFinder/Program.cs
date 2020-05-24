using FluentNHibernate.Cfg;
using FluentNHibernate.Cfg.Db;
using NHibernate;
using NHibernate.Tool.hbm2ddl;
using NHibernate.Cfg;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using OfficeOpenXml.Style;

namespace CollisionFinder
{



    class Program
    {
        private const string DbFile = "firstProgram.db";

        public static List<DB.Material> materialDB = new List<DB.Material>();
        public static List<DB.MaterialGroup> materialGroupDB = new List<DB.MaterialGroup>();
        public static List<DB.MaterialCode> materialCodeDB = new List<DB.MaterialCode>();
        public static List<DB.CustomHistory> customHistoryDB = new List<DB.CustomHistory>();

        static void Main(string[] args)
        {
            //string path = @"C:\Users\Alex\Desktop\файлы\ДВ ГПН Ямал 16.03.2020.zip\ДВ ГПН Ямал 16.03.2020_test.xlsx";
            //string newPath = @"C:\Users\Alex\Desktop\файлы\ДВ ГПН Ямал 16.03.2020.zip\ДВ ГПН Ямал 16.03.2020_test_2.xlsx";

            //string path_2 = @"C:\Users\Alex\Desktop\файлы\Выгрузка 509 17.03.2020.zip\Выгрузка 509 17.03.2020_test.xlsx";
            //string newPath_2 = @"C:\Users\Alex\Desktop\файлы\Выгрузка 509 17.03.2020.zip\Выгрузка 509 17.03.2020_test_2.xlsx";

            //string newPath_3 = @"C:\Users\Alex\Desktop\файлы\TotalFile\TotalFile.xlsx";
            //  MyCLI.Menu(); 
            Test(); // work this MTR CATALOG
           // Test2(); // Test features


        }
        static void Test2()
        {
            IerarhTemplate();
        }

        static void IerarhTemplate()
        {
            //int allCount = 1909;
            int allCount = 2304;
            string filepath = @"C:\Users\Alex\Desktop\nomenclatur4.xlsx";
            string filepath2 = @"C:\Users\Alex\Desktop\nomenclatur5.xlsx";
            using (var fs = new FileStream(filepath, FileMode.Open))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage doc = new ExcelPackage(fs))
                {
                    // Read excel sheets

                    Console.WriteLine("Чтение таблицы");

                    ExcelWorksheet sheet = doc.Workbook.Worksheets["Sheet1"];
                    int last_layer = 1;
                    for (int i = 6; i <= allCount; i++)
                    {
                        int c = (sheet.Cells[i, 1].Value ?? string.Empty).ToString().Count();
                        if (c != 0)
                        {
                            c--;
                            c /= 2;
                            last_layer = c;

                            sheet.Row(i).OutlineLevel = last_layer - 1;
                            sheet.Row(i).Collapsed = true;
                        }
                        else
                        {
                            sheet.Row(i).OutlineLevel = last_layer;
                            sheet.Row(i).Collapsed = true;
                        }
                    }
                    FileInfo fi = new FileInfo(filepath2);
                    doc.SaveAs(fi);
                    doc.Dispose();
                }

            }
        }

        static void Test()
        {

            /*
             * ВАЖНО
             * При добавлении новых файлов необходимо убедиться что первое по счету поле "Блокир." предназначенное для блокировки кода 
             * является не пустым!
             * В использованных раннее файлах было замено дублирование поля, столбец которого был пустым
             * АЛГОРИТМ ПОДБОРА БАЗОВОГО КОДА НЕ БУДЕТ РАБОТАТЬ КОРРЕКТНО И ВСЕ КОДЫ БУДУТ ПОМЕЧЕНЫ КАК ОБЕЗЛИЧЕННЫЕ
             */

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
            List<CodeCatalog> CodeCatalogList = new List<CodeCatalog>();
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet = doc.Workbook.Worksheets.Add("Sheet1");

                MTR_Catalog.ConvertEI(ref MaterialCatalog);

                FormBD(sheet, MaterialCatalog);

                


                CodeCatalogList = MTR_Catalog.Header(sheet, MaterialCatalog);

                for (int i = 1; i <= 26; i++)
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

        public static void FormGroup3Params(List<MTR_Catalog> MtrCatalogList)
        {
            var group = MtrCatalogList
                .Select(u => new
                {
                    groupName = u.GroupName,
                    groupCode = u.GroupCode,
                    groupCodeClass = u.NaimCodeClass
                }
                ).GroupBy(s => s.groupCode).Distinct();
            foreach (var t0 in group)
            {
                var ngroup = t0
                    .Distinct();
                CreateMaterialGroup3Params(ngroup);
            }
        }

        public static void FormGroup2Params(List<MTR_Catalog> MtrCatalogList)
        {
            var group = MtrCatalogList
                .Select(u => new
                {
                    groupCode = u.GroupCode,
                    groupCodeClass = u.NaimCodeClass
                }
                ).GroupBy(s => s.groupCode).Distinct();
            foreach (var t0 in group)
            {
                var ngroup = t0
                    .Distinct();
                CreateMaterialGroup2Params(ngroup);
            }
        }

        public static void FormBD(ExcelWorksheet sheet, List<MTR_Catalog> MtrCatalogList)
        {
            int MaterialIDCount = 0;
            var CodeCatalogList = new List<CodeCatalog>();

            //var group = MtrCatalogList
            //    .Select(u => new
            //    {
            //        groupName = u.GroupName,
            //        groupCode = u.GroupCode,
            //        groupCodeClass = u.NaimCodeClass
            //    }
            //    ).GroupBy(s => s.groupCode).Distinct();
            //foreach(var t0 in group)
            //{
            //    var ngroup = t0
            //        .Distinct();
            //    CreateMaterialGroup(ngroup);
            //}

            FormGroup2Params(MtrCatalogList);

            var ShortNameGroup = MtrCatalogList
        .GroupBy(s => s.MaterialFullName);
            foreach (var s0 in ShortNameGroup)
            {
                var NameGroup = s0
                .GroupBy(s => s.MaterialName);
                foreach (var s1 in NameGroup)
                {
                    DB.Material _material = new DB.Material();

                    //_material.CustomHistory = new List<DB.Custom_history>();
                    //_material.MaterialCode = new List<DB.Material_code>();
                    //_material.ID = MaterialIDCount; // for NHibernate
                    MaterialIDCount++;
                    _material.Material_fullname = s0.Key;
                    _material.Material_name = s1.Key;                    

                    foreach (var s2 in s1)
                    {
                        var FindGroup = materialGroupDB
                            .Where(u => u.Group_class_name == s2.NaimCodeClass)
                            .Where(u => u.Group_code == s2.GroupCode)
                            //.Where(u => u.Group_name == s2.GroupName) // uncomment for 3 params
                            .ToList();
                        FindGroup[0].Material.Add(_material);
                        _material.MaterialGroup = FindGroup[0];
                        break;
                    }

                    List<BaseCodeAtribute> baseCodeAtributeList = new List<BaseCodeAtribute>();
                    var pp = s1
                        .Select(u => new
                        {
                            code = u.MaterialCode,
                            date = u.DeliveryDate,
                            blockCode = u.BlockCode,
                            MU = u.BasisMU,
                            Brutto = u.Brutto
                            
                        });

                    foreach (var s in pp)
                    {
                        BaseCodeAtribute BCL = new BaseCodeAtribute();

                        var tp = Cast(s, new { code = "", date = "", blockCode = "", MU = "", Brutto = "" });
                        BCL.code = tp.code;
                        BCL.date = tp.date;
                        BCL.blockCode = tp.blockCode;
                        BCL.MU = tp.MU;
                        BCL.Brutto = tp.Brutto;
                        baseCodeAtributeList.Add(BCL);
                    }

                    // формируются коды
                    var difCode = s1
                        .GroupBy(x => x.MaterialCode)
                        .Select(x => x.First())
                        .Select(x => x.MaterialCode)
                        .ToList();
                    CodeCatalog cc = new CodeCatalog();
                    BaseCodeAtribute tmp_cc = new BaseCodeAtribute();
                    cc.Name = s1.Key;

                    tmp_cc = Functions.FindBaseCode(baseCodeAtributeList);
                    cc.BaseCode = tmp_cc.code;
                    cc.BaseBrutto = tmp_cc.Brutto;
                    cc.BaseMU = tmp_cc.MU;

                    cc.AltCode = difCode;
                    _material.Basic_code = cc.BaseCode;
                    _material.Measure_unit = cc.BaseMU;
                    _material.Brutto = cc.BaseBrutto;
                    CodeCatalogList.Add(cc);
                    CreateMaterialCode(ref _material, cc);

                    // формируется история
                    var gg = s1
                        .OrderBy(s => s.MaterialCode);
                    foreach (var s2 in gg)
                    {                       
                        CreateCustomHistory(ref _material, s2);
                    }
                    materialDB.Add(_material);
                }
            }
            //NHibernateWork();
            BDExcelOutput(); //вывод базы в excel
            FormGeneralMtrCatalogExcel(MtrCatalogList);
        }

        
        static void FormGeneralMtrCatalogExcel(List<MTR_Catalog> catalogs)
        {
            int allCount = 2304;
            int allNumSheet2 = 6;
            string filepath = @"C:\Users\Alex\Desktop\nomenclatur5.xlsx";
            string filepath2 = @"C:\Users\Alex\Desktop\nomenclatur2.xlsx";
            Console.WriteLine("Чтение документа: " + Functions.FirstNameFile(filepath));
            using (var fs = new FileStream(filepath, FileMode.Open))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage doc = new ExcelPackage(fs))
                {
                    // Read excel sheets

                    Console.WriteLine("Чтение таблицы");

                    ExcelWorksheet sheet = doc.Workbook.Worksheets["Sheet1"];
                    ExcelWorksheet sheet2 = doc.Workbook.Worksheets["Sheet2"];
                    sheet.Cells[1, 1, 5, 10].Copy(sheet2.Cells[1, 1, 5, 10]);
                    Console.WriteLine("Конец чтения таблицы");

                    //work with MTR
                    int aaa = 0;
                    foreach(var MMM in materialGroupDB)
                    {
                        int NumG = -1;
                        for (int i = 1; i <= allCount; i++)
                        {

                            if ((sheet.Cells[i, 1].Value ?? string.Empty).ToString() == MMM.Group_code)
                            {
                                NumG = i;                               
                                break;
                            }
                        }

                        if (NumG != -1)
                        {
                            var sum = MMM.Material.Count();
                            
                            sheet.InsertRow(NumG + 1, sum);
                            allCount += sum;
                            NumG++;
                            foreach (var s1 in MMM.Material)
                            {if(s1.Basic_code == "NONE")
                                {
                                    sheet.DeleteRow(NumG, 1, true);
                                    allCount--;
                                }
                                else
                                {
                                    sheet.Cells[NumG, 3].Value = s1.Basic_code;
                                    sheet.Cells[NumG, 4].Value = s1.Material_name;
                                    sheet.Cells[NumG, 5].Value = s1.Material_fullname;
                                    sheet.Cells[NumG, 5].Style.WrapText = true;
                                    sheet.Cells[NumG, 6].Value = s1.Measure_unit;
                                    sheet.Cells[NumG, 7].Value = s1.Brutto;
                                    NumG++;
                                }
                                
                            }
                        }
                        else
                        {
                            int numSheet2 = allNumSheet2;
                            bool first = true;
                            var sum = MMM.Material.Count();
                            sheet2.InsertRow(numSheet2 + 1, sum - 1);
                            allNumSheet2 += sum - 1;
                            foreach (var s1 in MMM.Material)
                            {
                                if(first)
                                {
                                    sheet2.Cells[numSheet2, 1].Value = MMM.Group_code;
                                    sheet2.Cells[numSheet2, 2].Value = MMM.Group_class_name;
                                    first = false;
                                }
                                sheet2.Cells[numSheet2, 3].Value = s1.Basic_code;
                                sheet2.Cells[numSheet2, 4].Value = s1.Material_name;
                                sheet2.Cells[numSheet2, 5].Value = s1.Material_fullname;
                                sheet2.Cells[numSheet2, 5].Style.WrapText = true;

                                sheet2.Cells[numSheet2, 6].Value = s1.Measure_unit;
                                sheet2.Cells[numSheet2, 7].Value = s1.Brutto;
                                numSheet2++;
                            }
                        }
                        
                    }

                    //var MMM = catalogs
                    //    .GroupBy(s => s.GroupCode);
                    //foreach (var s in MMM)
                    //{
                    //    int NumG = -1;
                    //    for (int i = 1; i <= 10000; i++)
                    //    {

                    //        if ((sheet.Cells[i, 1].Value ?? string.Empty).ToString() == s.Key)
                    //        {
                    //            NumG = i;
                    //            break;
                    //        }
                    //    }
                    //    if (NumG != -1)
                    //    {
                    //        var sum = s
                    //        .Count();
                    //        sheet.InsertRow(NumG + 1, sum - 1);
                    //        foreach (var s1 in s)
                    //        {
                    //            sheet.Cells[NumG, 4].Value = s1.MaterialName.ToString();
                    //            sheet.Cells[NumG, 5].Value = s1.MaterialFullName.ToString();
                    //            sheet.Cells[NumG, 6].Value = s1.BasisMU.ToString();
                    //            sheet.Cells[NumG, 7].Value = s1.Brutto.ToString();
                    //            NumG++;
                    //        }
                    //    }

                    //}


                    // save and close excel files
                    int minSize = 75;
                    int maxSize = 150;
                    for (int i = 4; i <= 5; i++)
                    {
                        if(i == 5)
                        {
                            sheet.Column(i).AutoFit(minSize, maxSize);
                        }
                        else
                        {
                            sheet.Column(i).AutoFit();
                        }
                        
                        sheet.Column(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                    for (int i = 1; i <= 5; i++)
                    {
                        if (i == 5)
                        {
                            sheet2.Column(i).AutoFit(minSize, maxSize);
                        }
                        else
                        {
                            sheet2.Column(i).AutoFit();
                        }

                        sheet2.Column(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                    //int j;
                    int last_layer = 1;
                    for (int i = 6; i <= allCount; i++)
                    {
                        int c = (sheet.Cells[i, 1].Value ?? string.Empty).ToString().Count();
                        if (c != 0)
                        {
                            c--;
                            c /= 2;
                            last_layer = c;

                            sheet.Row(i).OutlineLevel = last_layer - 1;
                            sheet.Row(i).Collapsed = true;
                        }
                        else
                        {
                            sheet.Row(i).OutlineLevel = last_layer;
                            sheet.Row(i).Collapsed = true;
                        }
                    }

                    // Assign borders
                    sheet2.Cells[1, 1, allNumSheet2, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet2.Cells[1, 1, allNumSheet2, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet2.Cells[1, 1, allNumSheet2, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet2.Cells[1, 1, allNumSheet2, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    // Assign borders
                    sheet.Cells[1, 1, allCount, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[1, 1, allCount, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[1, 1, allCount, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[1, 1, allCount, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                    FileInfo fi = new FileInfo(filepath2);
                    doc.SaveAs(fi);
                    doc.Dispose();
                }
            }
        }

        static void BDExcelOutput()
        {
            string fileName = @"C:\Users\Alex\Desktop\БД.xlsx";
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet1 = doc.Workbook.Worksheets.Add("Material");
                int n = 1;
                foreach (var s in materialDB)
                {
                    sheet1.Cells[n, 1].Value = s.ID;
                    sheet1.Cells[n, 2].Value = s.Basic_code;
                    sheet1.Cells[n, 3].Value = s.IsHide;
                    sheet1.Cells[n, 4].Value = s.Material_name;
                    sheet1.Cells[n, 5].Value = s.Material_fullname;
                    sheet1.Cells[n, 6].Value = "";
                    n++;
                }
                ExcelWorksheet sheet2 = doc.Workbook.Worksheets.Add("History");
                n = 1;
                foreach (var s in customHistoryDB)
                {
                    sheet2.Cells[n, 1].Value = s.ID;
                    //sheet2.Cells[n, 2].Value = s.Material_ID;
                    sheet2.Cells[n, 3].Value = s.Shipment_date;
                    sheet2.Cells[n, 4].Value = s.Consignee_detail;
                    sheet2.Cells[n, 5].Value = s.Basis_measure_unit;
                    sheet2.Cells[n, 6].Value = s.Count_BMU;
                    sheet2.Cells[n, 7].Value = s.Shipment_price_BMU;
                    sheet2.Cells[n, 8].Value = s.Alt_measure_unit;
                    sheet2.Cells[n, 9].Value = s.Count_AMU;
                    sheet2.Cells[n, 10].Value = s.Shipment_price_AMU;
                    n++;
                }
                ExcelWorksheet sheet3 = doc.Workbook.Worksheets.Add("Code");
                n = 1;
                foreach (var s in materialCodeDB)
                {
                    sheet3.Cells[n, 1].Value = s.ID;
                    //sheet3.Cells[n, 2].Value = s.Material_ID;
                    sheet3.Cells[n, 3].Value = s.Alternative_code;
                    sheet3.Cells[n, 4].Value = s.Basic_code;
                    n++;
                }
                ExcelWorksheet sheet4 = doc.Workbook.Worksheets.Add("MGroup");
                n = 1;
                foreach (var s in materialGroupDB)
                {
                    sheet4.Cells[n, 1].Value = s.ID;
                    sheet4.Cells[n, 2].Value = s.Group_name;
                    sheet4.Cells[n, 3].Value = s.Group_code;
                    sheet4.Cells[n, 4].Value = s.Group_class_name;
                    n++;
                }
                FileInfo fi = new FileInfo(fileName);

                doc.SaveAs(fi);
                doc.Dispose();
            }
        }

        static void GroupPrepareBD(List<MTR_Catalog> catalogs)
        {
            for (int i = 0; i < materialGroupDB.Count; i++)
            {
                var m = catalogs
                    .Where(u => u.GroupName == materialGroupDB[i].Group_name)
                    .Where(u => u.GroupCode == materialGroupDB[i].Group_code)
                    .Where(u => u.NaimCodeClass == materialGroupDB[i].Group_class_name)
                    .ToList();
                for (int j = 0; j < m.Count; j++)
                {
                    DB.Material material = new DB.Material();
                }
            }
        }

        static T Cast<T>(object obj, T type)
        {
            return (T)obj;
        }

        private static void CreateMaterialGroup3Params(IEnumerable<object> ngroup)
        {
            foreach (var s in ngroup)
            {
                DB.MaterialGroup _Group = new DB.MaterialGroup();
                //_Group.Material = new List<DB.Material>();
                //_Group.ID = materialGroupDB.Count(); // for NHibernate
                var tt = Cast(s, new { groupName = "", groupCode = "", groupCodeClass = "" });
                _Group.Group_name = tt.groupName;
                _Group.Group_code = tt.groupCode;
                _Group.Group_class_name = tt.groupCodeClass;
                materialGroupDB.Add(_Group);
            }
        }

        private static void CreateMaterialGroup2Params(IEnumerable<object> ngroup)
        {
            foreach (var s in ngroup)
            {
                DB.MaterialGroup _Group = new DB.MaterialGroup();
                //_Group.Material = new List<DB.Material>();
                //_Group.ID = materialGroupDB.Count(); // for NHibernate
                var tt = Cast(s, new { groupCode = "", groupCodeClass = "" });
                _Group.Group_code = tt.groupCode;
                _Group.Group_class_name = tt.groupCodeClass;
                materialGroupDB.Add(_Group);
            }
        }

        public static void CreateMaterialCode(ref DB.Material material, CodeCatalog codeCatalog)
        {
            //material.MaterialCode = new List<DB.Material_code>();
            foreach (var s in codeCatalog.AltCode)
            {
                DB.MaterialCode _Code = new DB.MaterialCode();
                //_Code.Material_ID = material.ID;
                _Code.Basic_code = codeCatalog.BaseCode;
                _Code.Alternative_code = s;
                //_Code.ID = materialCodeDB.Count; //for NHibernate
                _Code.Material = material; // new edit
                materialCodeDB.Add(_Code);

                // обратная привязка
                material.MaterialCode.Add(_Code);
            }
        }

        public static void CreateCustomHistory(ref DB.Material material, MTR_Catalog catalog)
        {
            //material.CustomHistory = new List<DB.Custom_history>();

            double countA, countB, priceA, priceB;
            DB.CustomHistory _History = new DB.CustomHistory();
            //_History.Material_ID = material.ID;
            _History.Shipment_date = catalog.DeliveryDate;
            _History.Consignee_detail = catalog.ConsigneeDetail;

            _History.Basis_measure_unit = catalog.BasisMU;
            Double.TryParse(catalog.BasisMUCount, out countB);
            _History.Count_BMU = countB;
            Double.TryParse(catalog.BasisMUPrice, out priceB);
            _History.Shipment_price_BMU = priceB;

            _History.Alt_measure_unit = catalog.AltMU;
            Double.TryParse(catalog.AltMUCount, out countA);
            _History.Count_AMU = countA;
            Double.TryParse(catalog.AltMUPrice, out priceA);
            _History.Shipment_price_AMU = priceA;

            _History.Material = material;
            //_History.ID = customHistoryDB.Count; // for NHibernate
            customHistoryDB.Add(_History);
            material.CustomHistory.Add(_History);

        }

        static void NHibernateWork()
        {
            var sessionFactory = CreateSessionFactory();
            using (var session = sessionFactory.OpenSession())
            {
                using (var transaction = session.BeginTransaction())
                {
                    foreach (var obj in materialGroupDB)
                    {
                        session.SaveOrUpdate(obj);
                    }
                    transaction.Commit();
                }
            }
        }

        private static ISessionFactory CreateSessionFactory()
        {
            var connectionStr = "Server=127.0.0.1;Port=5432;Database=MtrCatalog;User Id=postgres;Password=123456;";
            return Fluently.Configure()
              .Database(
                PostgreSQLConfiguration.Standard.ConnectionString(connectionStr))
              .Mappings(m => m.FluentMappings.AddFromAssemblyOf<Program>())
              //.ExposeConfiguration(cfg => { new SchemaExport(cfg).Create(false, true);})
              .ExposeConfiguration(BuildSchema)
              .BuildSessionFactory();
        }

        private static void BuildSchema(NHibernate.Cfg.Configuration config)
        {

            // delete the existing db on each run
            if (File.Exists(DbFile))
                File.Delete(DbFile);

            // this NHibernate tool takes a configuration (with mapping info in)
            // and exports a database schema from it
            new SchemaExport(config)
              .Create(false, true);

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



