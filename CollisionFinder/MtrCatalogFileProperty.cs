using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollisionFinder
{

    class MtrCatalogFileProperty
    {
        private readonly int _find_size = 500;
        public string FilePath { get; set; }

        public int FirstRow { get; set; } 

        public uint LastRow { get; set; }
    
        // CODE

        /// <summary>
        /// Код МТР
        /// </summary>
        public int MaterialCodeCol { get; set; }

        /// <summary>
        /// Проверка блокировки кода МТР
        /// </summary>
        public int BlockCodeCol { get; set; }

        // MTR NAME

        /// <summary>
        /// Краткое наименование МТР
        /// </summary>
        public int MaterialNameCol { get; set; }

        /// <summary>
        /// Полное наименование МТР
        /// </summary>
        public int MaterialFullName1Col { get; set; }

        public int MaterialFullName2Col { get; set; }

        // STRUCTURE POINT

        /// <summary>
        /// Название группы МТР
        /// </summary>
        public int GroupNameCol { get; set; }

        /// <summary>
        /// Код класса МТР
        /// </summary>
        public int GroupCodeCol { get; set; }

        // MTR PRICE

        /// <summary>
        /// Наим. Код класс
        /// </summary>
        public int NaimCodeClassCol { get; set; } // Temp name!

        /// <summary>
        /// Реквизиты грузополучателя
        /// </summary>
        public int ConsigneeDetailCol { get; set; }

        /// <summary>
        /// Дата поставки
        /// </summary>
        public int DeliveryDateCol { get; set; }

        /// <summary>
        /// Базисная ЕИ
        /// </summary>
        public int BasisMUCol { get; set; }

        /// <summary>
        /// Кол-во к закупу, БЕИ
        /// </summary>
        public int BasisMUCountCol { get; set; }

        /// <summary>
        /// Цена поставки с НДС, БЕИ
        /// </summary>
        public int BasisMUPriceCol { get; set; }

        /// <summary>
        /// Альтернативная ЕИ
        /// </summary>
        public int AltMUCol { get; set; }

        /// <summary>
        ///  Кол-во к закупу, АЕИ
        /// </summary>
        public int AltMUCountCol { get; set; }

        /// <summary>
        /// Цена поставки с НДС, АЕИ
        /// </summary>
        public int AltMUPriceCol { get; set; }

        public int SPPNameCol { get; set; }

        public int SPPElemCol { get; set; }

        public int OKPD2Col { get; set; }

        public int OKPD2CodeCol { get; set; }

        public int BruttoCol { get; set; }

        public int Kol_voSCHFCol { get; set; }

        public int SumSCHFWithoutNDSCol { get; set; }

        public int DateSchfCol { get; set; }

        public MtrCatalogFileProperty FindColumns(ref ExcelWorksheet sheet, MtrCatalogFileProperty tt)
        {
            var ans = tt;

            for(int i = 1; i <= _find_size; i++)
            {
                if(sheet.Cells[1,i].Value.ToString() == "Материал")
                {
                    ans.MaterialCodeCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Блокир.")
                {
                    ans.BlockCodeCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Материал Имя")
                {
                    ans.MaterialNameCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Материал имя (полное)1")
                {
                    ans.MaterialFullName1Col = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Материал имя (полное)2")
                {
                    ans.MaterialFullName2Col = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Название группы")
                {
                    ans.GroupNameCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Код класса МТР")
                {
                    ans.GroupCodeCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Наим.Код кл.")
                {
                    ans.NaimCodeClassCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Рекв.Грузополучателя")
                {
                    ans.ConsigneeDetailCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Срок поставки")
                {
                    ans.DeliveryDateCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Базисная ЕИ")
                {
                    ans.BasisMUCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Кол-во к закупу, БЕИ")
                {
                    ans.BasisMUCountCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Цена поставки с НДС")
                {
                    ans.BasisMUPriceCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "АЕИ заказа")
                {
                    ans.AltMUCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Кол-во к закупу, АЕИ")
                {
                    ans.AltMUCountCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Цена поставки с НДС за АЕИ")
                {
                    ans.AltMUPriceCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "СПП имя")
                {
                    ans.SPPNameCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "СПП элемент")
                {
                    ans.SPPElemCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "ОКПД2")
                {
                    ans.OKPD2Col = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Код по ОКПД2")
                {
                    ans.OKPD2CodeCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Вес брутто")
                {
                    ans.BruttoCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Количество по Сч/ф")
                {
                    ans.Kol_voSCHFCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Сумма по Сч/ф без НДС")
                {
                    ans.SumSCHFWithoutNDSCol = i;
                    break;
                }
            }

            for (int i = 1; i <= _find_size; i++)
            {
                if (sheet.Cells[1, i].Value.ToString() == "Дата Сч/ф")
                {
                    ans.DateSchfCol = i;
                    break;
                }
            }
            return ans;
        }        
    }
}