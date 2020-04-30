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

        public static void DateCompire()
        {

        }
    }
}
