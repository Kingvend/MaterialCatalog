using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CollisionFinder
{
    static class Functions
    {
        

        /// <summary>
        /// Выделяет имя файла из пути
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string FirstNameFile(string path)
        {
            string s = "";
            for (int i = path.Length - 1; i >= 0; i--)
            {
                if (path[i] != '\\')
                {
                    s = s.Insert(0, path[i].ToString());
                }
                else
                {
                    break;
                }

            }
            return s;
        }

        public static string BlockCodeConvert(string str)
        {
            string ans = "";
            switch (str)
            {
                case "Нет": ans = "0"; break;
                case "Да": ans = "1"; break;
            }
            return ans;
        }

        /// <summary>
        /// преобразует индекс колонки excel в порядковый номер
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int ConvertNumberColumnInExcel(string s)
        {
            int ans = 0;
            int pow = 0;
            for (int i = s.Length - 1; i >= 0; i--)
            {
                ans += (s[i] - 'A' + 1) * (int)Math.Pow(26, pow);
                pow++;
            }
            return ans;
        }
        /// <summary>
        /// Преобразование строки
        /// </summary>
        /// <param name="NewString"></param>
        /// <returns></returns>
        public static string ChangeString(string NewString)
        {
            string tmp_string = "";
            tmp_string = NewString;
            if (tmp_string == null)
            {
                tmp_string = "";
            }
            tmp_string = tmp_string.ToUpper();

            // удаление лишних символов
            for (int i = 33; i <= 47; i++)
            {
                tmp_string = tmp_string.Replace((char)i, (char)32);

            }
            tmp_string = tmp_string.Replace(" ", "");

            return tmp_string;
        }
        /// <summary>
        /// нахождение  
        /// </summary>

        public static int DateCompire(string s1, string s2)
        {

            // 0 - s1 > s2; 1 - s1 < s2; 2 - s1 = s2;

            int day1, month1, year1;
            int day2, month2, year2;
            //int hour1, minute1, second1;
            //int hour2, minute2, second2;

            day1 = Int32.Parse(s1.Substring(0, 2));
            month1 = Int32.Parse(s1.Substring(3, 2));
            year1 = Int32.Parse(s1.Substring(6, 4));

            day2 = Int32.Parse(s2.Substring(0, 2));
            month2 = Int32.Parse(s2.Substring(3, 2));
            year2 = Int32.Parse(s2.Substring(6, 4));

            if (year1 == year2)
            {
                if (month1 == month2)
                {
                    //if (day1 == day2) // only year and month
                    //{
                    //    return 2;
                    //}
                    //else
                    //{
                    //    return day1 < day2 ? 1 : 0;
                    //}
                    return 2;
                }
                else
                {
                    return month1 < month2 ? 1 : 0;
                }
            }
            else
            {
                return year1 < year2 ? 1 : 0;
            }
        }

        public static string FindBaseCode(List<BaseCodeAtribute> baseCodeAtributes0)
        {
            var baseCodeAtributes = baseCodeAtributes0
                .Where(s => s.blockCode == "0");
            List<BaseCodeAtribute> baseCodeAtributesBestDate = new List<BaseCodeAtribute>();
            string LastDate = "00.00.0000 0:00:00";
            foreach (var d in baseCodeAtributes)
            {
                if (DateCompire(d.date, LastDate) == 0)
                {
                    baseCodeAtributesBestDate.Clear();
                    baseCodeAtributesBestDate.Add(d);
                    LastDate = d.date;
                }
                else
                {
                    if (DateCompire(d.date, LastDate) == 2)
                    {
                        baseCodeAtributesBestDate.Add(d);
                        LastDate = d.date;
                    }
                }
            }

            if (baseCodeAtributesBestDate.Count > 1)
            {
                //var tt = baseCodeAtributesBestDate // for test
                //    .Select(s => s.code)
                //    .Distinct()
                //    .Count();
                //if (tt > 1)
                //{
                //    int i = 2;
                //}

                // Старый способ
                //var unic = baseCodeAtributesBestDate
                //    .GroupBy(s => s.code)
                //   .OrderByDescending(s => s.Count())
                //   .First()
                //   .Key;

                var Max = baseCodeAtributesBestDate
                    .GroupBy(s => s.code)
                   .Max(s => s.Count());

                var OftenCode = baseCodeAtributesBestDate
                   .GroupBy(s => s.code)
                   .OrderByDescending(s => s.Count())
                   .TakeWhile(x => x.Count() == Max)
                   .ToList();
                
                var MaxCode = OftenCode
                   .OrderByDescending(s => s.Key)
                   .First()
                   .Key;
                return MaxCode;
            }
            else
            {
                if (baseCodeAtributesBestDate.Count == 0)
                {
                    return "NONE";
                }
                return baseCodeAtributesBestDate[0].code;
            }
        }

        static List<MTR_Catalog> CatalogForYear(List<MTR_Catalog> catalogs, int year)
        {
            var catalog = new List<MTR_Catalog>();
            int year1;

            // code
            foreach (var c in catalogs)
            {
                //month1 = Int32.Parse(c.DeliveryDate.Substring(3, 2));
                year1 = Int32.Parse(c.DeliveryDate.Substring(6, 4));
                if (year1 == year)
                {
                    catalog.Add(c);
                }
            }
            return catalog;
        }

        public static double SumMtr(List<MTR_Catalog> catalogs, double Koef1, double Koef2, int CurentYear)
        {
            var Catalog1 = new List<MTR_Catalog>();
            var Catalog2 = new List<MTR_Catalog>();
            var Catalog3 = new List<MTR_Catalog>();
            NumberFormatInfo provider = new NumberFormatInfo();
            provider.NumberDecimalSeparator = ",";

            double Sum1 = 0, Sum2 = 0, Sum3 = 0;
            double V1 = 0, V2 = 0, V3 = 0;

            double ans = 0.0;
            // code
            // 3 catalogs for last 3 year
            Catalog1.AddRange(CatalogForYear(catalogs, CurentYear));
            Catalog2.AddRange(CatalogForYear(catalogs, CurentYear - 1));
            Catalog3.AddRange(CatalogForYear(catalogs, CurentYear - 2));

            foreach (var s in Catalog1)
            {
                double tmpSum, tmpV;
                tmpSum = Convert.ToDouble(s.SumSCHFWithoutNDS, provider);
                Sum1 += tmpSum;
                tmpV = Convert.ToDouble(s.Kol_voSCHF, provider);
                V1 += tmpV;
            }

            foreach (var s in Catalog2)
            {
                double tmpSum, tmpV;
                tmpSum = Convert.ToDouble(s.SumSCHFWithoutNDS, provider);
                Sum2 += tmpSum;
                tmpV = Convert.ToDouble(s.Kol_voSCHF, provider);
                V2 += tmpV;
            }

            foreach (var s in Catalog3)
            {
                double tmpSum, tmpV;
                tmpSum = Convert.ToDouble(s.SumSCHFWithoutNDS, provider);
                Sum3 += tmpSum;
                tmpV = Convert.ToDouble(s.Kol_voSCHF, provider);
                V3 += tmpV;
            }

            double TotalV = V1 + V2 + V3;
            double Sr1, Sr2, Sr3;
            if (V1 == 0)
            {
                Sr1 = 0;
            }
            else
            {
                Sr1 = Sum1 / V1;
            }

            if (V2 == 0)
            {
                Sr2 = 0;
            }
            else
            {
                Sr2 = Sum2 / V2 * Koef1;
            }

            if (V3 == 0)
            {
                Sr3 = 0;
            }
            else
            {
                Sr3 = Sum3 / V3 * Koef1 * Koef2;
            }

            if (TotalV == 0)
            {
                ans = 0;
            }
            else
            {
                ans = (V1 / TotalV * Sr1) + (V2 / TotalV * Sr2) + (V3 / TotalV * Sr3);
            }

            return ans;
        }
    }
}
