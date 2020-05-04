using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder
{
    class MTR_Catalog
    {
        // CODE
        

        /// <summary>
        /// Код МТР
        /// </summary>
        public string MaterialCode { get; set; }

        /// <summary>
        /// Проверка блокировки кода МТР
        /// </summary>
        public string BlockCode { get; set; }

        // MTR NAME

        /// <summary>
        /// Краткое наименование МТР
        /// </summary>
        public string MaterialName { get; set; }

        /// <summary>
        /// Полное наименование МТР
        /// </summary>
        public string MaterialFullName { get; set; }

        // STRUCTURE POINT

        /// <summary>
        /// Название группы МТР
        /// </summary>
        public string GroupName { get; set; }

        /// <summary>
        /// Код класса МТР
        /// </summary>
        public string GroupCode { get; set; }

        // MTR PRICE

        /// <summary>
        /// Наим. Код класс
        /// </summary>
        public string NaimCodeClass { get; set; } // Temp name!

        /// <summary>
        /// Реквизиты грузополучателя
        /// </summary>
        public string ConsigneeDetail { get; set; }

        /// <summary>
        /// Дата поставки
        /// </summary>
        public string DeliveryDate { get; set; }

        /// <summary>
        /// Базисная ЕИ
        /// </summary>
        public string BasisMU { get; set; }

        /// <summary>
        /// Кол-во к закупу, БЕИ
        /// </summary>
        public string BasisMUCount { get; set; }

        /// <summary>
        /// Цена поставки с НДС, БЕИ
        /// </summary>
        public string BasisMUPrice { get; set; }

        /// <summary>
        /// Альтернативная ЕИ
        /// </summary>
        public string AltMU { get; set; }

        /// <summary>
        ///  Кол-во к закупу, АЕИ
        /// </summary>
        public string AltMUCount { get; set; }

        /// <summary>
        /// Цена поставки с НДС, АЕИ
        /// </summary>
        public string AltMUPrice { get; set; }

        public bool OrangeStatus { get; set; }

        public string SPPName { get; set; }

        public string SPPElem { get; set; }

        public string OKPD2 { get; set; }

        public string OKPD2Code { get; set; }

        public static List<CodeCatalog> Header(ExcelWorksheet sheet, List<MTR_Catalog> MtrCatalogList)
        {
            Dictionary<string, int> Full = new Dictionary<string, int>();
            Dictionary<string, bool> Short = new Dictionary<string, bool>();

            sheet.Cells[1, 1].Value = "СПРАВОЧНИК МТР";
            sheet.Cells[1, 1, 1, 8].Merge = true;

            sheet.Cells[1, 9].Value = "ЦЕНА МТР";
            sheet.Cells[1, 9, 1, 16].Merge = true;

            sheet.Cells[2, 1].Value = "Код МТР";
            sheet.Cells[2, 1, 2, 2].Merge = true;

            sheet.Cells[2, 3].Value = "Наименование МТР";
            sheet.Cells[2, 3, 2, 5].Merge = true;

            sheet.Cells[2, 6].Value = "Структурирование справочника";
            sheet.Cells[2, 6, 2, 8].Merge = true;

            sheet.Cells[2, 9].Value = "Базис поставки";

            sheet.Cells[2, 10].Value = "Актуальность цены";

            sheet.Cells[2, 11].Value = "В базовых ЕИ";
            sheet.Cells[2, 11, 2, 13].Merge = true;

            sheet.Cells[2, 14].Value = "В альтернативных ЕИ";
            sheet.Cells[2, 14, 2, 16].Merge = true;

            var prop = typeof(MtrCatalogFileProperty).GetProperties();
            for (int i = 3; i < prop.Length; i++)
            {
                sheet.Cells[3, i - 2].Value = prop[i].Name.ToString();
            }

            sheet.Cells[2, 17].Value = "Сумма по базисным ЕИ";
            sheet.Cells[2, 18].Value = "Сумма по альтернативным ЕИ";

            sheet.Cells[2, 19].Value = "СПП имя";
            sheet.Cells[2, 20].Value = "СПП код";
            sheet.Cells[2, 21].Value = "код по ОКПД2";
            sheet.Cells[2, 22].Value = "ОКПД2";

            //пометка данных с коллизией
            var numRow = 4;

            var ShortNameGroup = MtrCatalogList
                .GroupBy(s => s.MaterialFullName);
            foreach (var s0 in ShortNameGroup)
            {
                Full.Add(s0.Key, 0);
                var NameGroup = s0
                .GroupBy(s => s.MaterialName);
                numRow++;
                foreach (var s1 in NameGroup)
                {
                    if(!Short.ContainsKey(s1.Key))
                    Short.Add(s1.Key, false);
                    int header_1_Row = numRow;

                    //double sumB = 0;
                    //double sumA = 0;
                    //double countB, countA;
                    //double priceB, priceA;
                    //bool IsOrange = false;

                    //sheet.Cells[numRow, 3].Value = s1.Key.ToString();



                    var tt = s1
                       .Select(s => s.MaterialCode)
                       .Distinct()
                       .Count();
                    if (tt > 1)
                    {
                        //sheet.Cells[numRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //sheet.Cells[numRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                        Full[s0.Key]++;
                        Short[s1.Key] = true;
                    }
                    else
                    {
                        //sheet.Cells[numRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //sheet.Cells[numRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }
                }
            }
            //

            /*var*/ numRow = 4;
            var CodeCatalogList = new List<CodeCatalog>();

            /*var*/ ShortNameGroup = MtrCatalogList
                .GroupBy(s => s.MaterialFullName);
            foreach(var s0 in ShortNameGroup)
            {
                //if (Full[s0.Key] == 0) continue;
                sheet.Cells[numRow, 4].Value = s0.Key.ToString();
                sheet.Cells[numRow, 4].Style.WrapText = true;
                numRow++;

                var NameGroup = s0
                .GroupBy(s => s.MaterialName);              
                foreach (var s1 in NameGroup)
                {
                    var difCode = s1.GroupBy(x => x.MaterialCode).Select(x => x.First()).Select(x => x.MaterialCode).ToList();
                    //if(difCode.Count > 1 )
                    //{
                    //    int i = 2 + 2;
                    //}
                    CodeCatalog cc = new CodeCatalog();

                    cc.Name = s1.Key;
                    cc.BaseCode = "";
                    cc.AltCode = difCode;
                    CodeCatalogList.Add(cc);
                    int header_1_Row = numRow;
                  
                    double sumB = 0;
                    double sumA = 0;
                    double countB, countA;
                    double priceB, priceA;
                    //bool IsOrange = false;

                    //if (Short[s1.Key] == false) continue;
                    sheet.Cells[numRow, 3].Value = s1.Key.ToString();
                 
                    var tt = s1
                       .Select(s => s.MaterialCode)
                       .Distinct()
                       .Count();
                    if (tt > 1)
                    {
                        sheet.Cells[numRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[numRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                    }
                    else
                    {
                        //continue;
                        sheet.Cells[numRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sheet.Cells[numRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }
                                           
                    var gg = s1
                        .OrderBy(s => s.MaterialCode);
                    foreach (var s2 in gg)
                    {
                        sheet.Cells[numRow, 1].Value = s2.MaterialCode;
                        sheet.Cells[numRow, 2].Value = s2.BlockCode;
                        //sheet.Cells[numRow, 3].Value = s2.MaterialName;
                        //sheet.Cells[numRow, 4].Value = s2.MaterialFullName;                       
                        //sheet.Cells[numRow, 4].AutoFitColumns(minSize, maxSize);
                        sheet.Cells[numRow, 6].Value = s2.GroupName;
                        sheet.Cells[numRow, 7].Value = s2.GroupCode;
                        sheet.Cells[numRow, 8].Value = s2.NaimCodeClass;
                        sheet.Cells[numRow, 9].Value = s2.ConsigneeDetail;
                        sheet.Cells[numRow, 10].Value = s2.DeliveryDate;
                        sheet.Cells[numRow, 11].Value = s2.BasisMU;
                        sheet.Cells[numRow, 12].Value = s2.BasisMUCount;
                        sheet.Cells[numRow, 13].Value = s2.BasisMUPrice;

                        if (Double.TryParse(s2.BasisMUCount, out countB) && Double.TryParse(s2.BasisMUPrice, out priceB))
                        {
                            sumB += (countB * priceB);
                        }
                        sheet.Cells[numRow, 14].Value = s2.AltMU;
                        sheet.Cells[numRow, 15].Value = s2.AltMUCount;
                        sheet.Cells[numRow, 16].Value = s2.AltMUPrice;

                        sheet.Cells[numRow, 19].Value = s2.SPPName;
                        sheet.Cells[numRow, 20].Value = s2.SPPElem;
                        sheet.Cells[numRow, 21].Value = s2.OKPD2Code;
                        sheet.Cells[numRow, 22].Value = s2.OKPD2;

                        if (Double.TryParse(s2.AltMUCount, out countA) && Double.TryParse(s2.AltMUPrice, out priceA))
                        {
                            sumA += (countA * priceA);
                        }
                       
                        numRow++;
                    }
                    sheet.Cells[header_1_Row, 17].Value = sumB.ToString();
                    sheet.Cells[header_1_Row, 18].Value = sumA.ToString();
                }
            }

            for(int i = 4; i < numRow; i++)
            {
                if(sheet.Cells[i, 4].Value == null)
                {
                    if (i == 4) continue;
                    else
                    {
                        if(sheet.Cells[i, 3].Value == null)
                        {
                            sheet.Row(i).OutlineLevel = 2;
                            sheet.Row(i).Collapsed = true;
                        }
                        else
                        {
                            sheet.Row(i).OutlineLevel = 1;
                            sheet.Row(i).Collapsed = true;
                        }                      
                    }
                }
            }
            return CodeCatalogList;
        }
    }
}
