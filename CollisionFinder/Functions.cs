using System;
using System.Collections.Generic;
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
               if(path[i] != '\\')
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
        /// <summary>
        /// преобразует индекс колонки excel в порядковый номер
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static int ConvertNumberColumnInExcel(string s)
        {
            int ans = 0;
            int pow = 0;
            for(int i = s.Length-1;i >=0;i--)
            {
                ans += (s[i] - 'A' + 1) * (int)Math.Pow(26,pow);
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
            if(tmp_string == null)
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
        public static void FindHeader(string FilePath, int FirstRow, int LastRow, string[] Headers)
        {

        }

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

            if(year1 == year2)
            {
                if(month1 == month2)
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

        public static string FindBaseCode(List<BaseCodeAtribute> baseCodeAtributes)
        {
            List<BaseCodeAtribute> baseCodeAtributesBestDate = new List<BaseCodeAtribute>();
            string LastDate = "00.00.0000 0:00:00";
            foreach(var d in baseCodeAtributes)
            {
                if(DateCompire(d.date,LastDate) == 0)
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

            if(baseCodeAtributesBestDate.Count > 1)
            {
                //var tt = baseCodeAtributesBestDate // for test
                //    .Select(s => s.code)
                //    .Distinct()
                //    .Count();
                //if (tt > 1)
                //{
                //    int i = 2;
                //}
                var unic = baseCodeAtributesBestDate
                    .GroupBy(s => s.code)
                   .OrderByDescending(s => s.Count())
                   .First()
                   .Key;
                return unic;

            }
            else
            {
                return baseCodeAtributesBestDate[0].code;
            }

            return baseCodeAtributesBestDate[0].code;
        }
    }
}
