using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace CollisionFinder
{
    /// <summary>
    /// Класс, сожержащий необходимую информацию по отчету 
    /// </summary>
    class Report
    {
        /// <summary>
        /// формирование отчета о коллизиях
        /// </summary>
        /// <param name="collision_1">коллизия связанная с кодом материала</param>
        /// <param name="collision_2">коллизия связанная с наименованием материала</param>
        /// <param name="firstRow">номер строки, с которой начнентся заполнение таблицы (зависит от шапки таблицы)</param>
        /// <param name="fileName">имя файла, в котором будет сформирован отчет</param>
        /// 
        public static void ReportGenerate(List<Collision> collision_1, List<Collision> collision_2, List<Collision> collision_3, List<Collision> collision_4, int firstRow, string fileName)
        {
            double minSize = 50;
            double maxSize = 150;
            using (ExcelPackage doc = new ExcelPackage())
            {
                ExcelWorksheet sheet = doc.Workbook.Worksheets.Add("Sheet1");
                int numberRow = firstRow;
                sheet.Cells[numberRow, 1].Value = "ОТЧЕТ";
                sheet.Cells[numberRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                sheet.Cells[numberRow, 1, numberRow, 5].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1].Style.Font.Bold = true;
                sheet.Cells[numberRow, 1, numberRow, 5].Merge = true;
                numberRow++;
                sheet.Cells[numberRow, 1].Value = "Параметр группировки";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 2].Merge = true;
                sheet.Cells[numberRow, 3].Value = "Код материала";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3, numberRow, 5].Merge = true;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;
                sheet.Cells[numberRow, 1].Value = "код материала";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 2].Value = "наименование (краткое)";
                sheet.Cells[numberRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3].Value = "наименование (полное)";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 4].Value = "имя исходного файла";
                sheet.Cells[numberRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 5].Value = "строка, №";
                sheet.Cells[numberRow, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;
                var GroupCodeMaterial = collision_1
                    .GroupBy(s => s.Code);
                foreach (var s1 in GroupCodeMaterial)
                {
                    sheet.Cells[numberRow, 1].Value = s1.Key.ToString();
                    sheet.Cells[numberRow, 1].AutoFitColumns();
                    foreach (var s2 in s1)
                    {
                        sheet.Cells[numberRow, 2].Value = s2.Name.ToString();
                        sheet.Cells[numberRow, 2].AutoFitColumns();
                        sheet.Cells[numberRow, 3].Value = s2.FullName.ToString();
                        sheet.Cells[numberRow, 3].AutoFitColumns(minSize, maxSize);
                        sheet.Cells[numberRow, 3].Style.WrapText = true;
                        sheet.Cells[numberRow, 4].Value = s2.FileSource;
                        sheet.Cells[numberRow, 4].AutoFitColumns();
                        sheet.Cells[numberRow, 5].Value = s2.RowNumber.ToString();
                        sheet.Cells[numberRow, 5].AutoFitColumns();
                        numberRow++;
                    }
                }

                numberRow++;
                sheet.Cells[numberRow, 1].Value = "Параметр группировки";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 2].Merge = true;
                sheet.Cells[numberRow, 3].Value = "Наименование материала";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3, numberRow, 5].Merge = true;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;
                sheet.Cells[numberRow, 2].Value = "код материала";
                sheet.Cells[numberRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1].Value = "наименование (краткое)";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3].Value = "наименование (полное)";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 4].Value = "имя исходного файла";
                sheet.Cells[numberRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 5].Value = "строка, №";
                sheet.Cells[numberRow, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;

                var GroupNameMaterial = collision_2
                    .GroupBy(s => s.Name);
                foreach (var s1 in GroupNameMaterial)
                {
                    sheet.Cells[numberRow, 1].Value = s1.Key.ToString();
                    sheet.Cells[numberRow, 1].AutoFitColumns();
                    foreach (var s2 in s1)
                    {
                        sheet.Cells[numberRow, 2].Value = s2.Code.ToString();
                        sheet.Cells[numberRow, 2].AutoFitColumns();
                        sheet.Cells[numberRow, 3].Value = s2.FullName.ToString();
                        sheet.Cells[numberRow, 3].AutoFitColumns(minSize, maxSize);
                        sheet.Cells[numberRow, 3].Style.WrapText = true;
                        sheet.Cells[numberRow, 4].Value = s2.FileSource;
                        sheet.Cells[numberRow, 4].AutoFitColumns();
                        sheet.Cells[numberRow, 5].Value = s2.RowNumber.ToString();
                        sheet.Cells[numberRow, 5].AutoFitColumns();
                        numberRow++;
                    }
                }

                numberRow++;
                sheet.Cells[numberRow, 1].Value = "Параметр группировки";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 2].Merge = true;
                sheet.Cells[numberRow, 3].Value = "Наименование материала";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3, numberRow, 5].Merge = true;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;
                sheet.Cells[numberRow, 2].Value = "код материала";
                sheet.Cells[numberRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1].Value = "наименование (краткое)";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3].Value = "наименование (полное)";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 4].Value = "имя исходного файла";
                sheet.Cells[numberRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 5].Value = "строка, №";
                sheet.Cells[numberRow, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;

                var GroupMessureMaterial = collision_3
                    .GroupBy(s => s.Name);
                foreach (var s1 in GroupMessureMaterial)
                {
                    sheet.Cells[numberRow, 1].Value = s1.Key.ToString();
                    sheet.Cells[numberRow, 1].AutoFitColumns();
                    foreach (var s2 in s1)
                    {
                        sheet.Cells[numberRow, 2].Value = s2.Code.ToString();
                        sheet.Cells[numberRow, 2].AutoFitColumns();
                        sheet.Cells[numberRow, 3].Value = s2.FullName.ToString();
                        sheet.Cells[numberRow, 3].AutoFitColumns(minSize, maxSize);
                        sheet.Cells[numberRow, 3].Style.WrapText = true;
                        sheet.Cells[numberRow, 4].Value = s2.FileSource;
                        sheet.Cells[numberRow, 4].AutoFitColumns();
                        sheet.Cells[numberRow, 5].Value = s2.RowNumber.ToString();
                        sheet.Cells[numberRow, 5].AutoFitColumns();
                        if(s2.Messure == null)
                        {
                            sheet.Cells[numberRow, 6].Value = "";
                        }
                        else
                        {
                            sheet.Cells[numberRow, 6].Value = (s2.Messure.ToString());
                        }
                        sheet.Cells[numberRow, 6].AutoFitColumns();
                        numberRow++;
                    }
                }

                numberRow++;
                sheet.Cells[numberRow, 1].Value = "Параметр группировки";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 2].Merge = true;
                sheet.Cells[numberRow, 3].Value = "Наименование материала";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3, numberRow, 5].Merge = true;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;
                sheet.Cells[numberRow, 2].Value = "код материала";
                sheet.Cells[numberRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1].Value = "наименование (краткое)";
                sheet.Cells[numberRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 3].Value = "наименование (полное)";
                sheet.Cells[numberRow, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 4].Value = "имя исходного файла";
                sheet.Cells[numberRow, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 5].Value = "строка, №";
                sheet.Cells[numberRow, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[numberRow, 1, numberRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                numberRow++;

                var GroupFullNameMaterial = collision_4
                    .GroupBy(s => s.FullName);
                foreach (var s1 in GroupFullNameMaterial)
                {
                    sheet.Cells[numberRow, 3].Value = s1.Key.ToString();
                    sheet.Cells[numberRow, 3].AutoFitColumns();
                    sheet.Cells[numberRow, 3].Style.WrapText = true;
                    foreach (var s2 in s1)
                    {
                        sheet.Cells[numberRow, 2].Value = s2.Code.ToString();
                        sheet.Cells[numberRow, 2].AutoFitColumns();
                        sheet.Cells[numberRow, 1].Value = s2.Name.ToString();
                        sheet.Cells[numberRow, 1].AutoFitColumns(minSize, maxSize);
                        sheet.Cells[numberRow, 1].Style.WrapText = true;
                        sheet.Cells[numberRow, 4].Value = s2.FileSource;
                        sheet.Cells[numberRow, 4].AutoFitColumns();
                        sheet.Cells[numberRow, 5].Value = s2.RowNumber.ToString();
                        sheet.Cells[numberRow, 5].AutoFitColumns();
                        numberRow++;
                    }
                }

                sheet.Column(1).AutoFit();
                sheet.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Column(2).AutoFit();
                sheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Column(4).AutoFit();
                sheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Column(5).AutoFit();
                sheet.Column(5).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                sheet.Column(6).AutoFit();
                sheet.Column(6).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                FileInfo fi = new FileInfo(fileName);

                    doc.SaveAs(fi);             
                doc.Dispose();
            }
        }
    }

}



