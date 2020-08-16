using System.IO;
using OfficeOpenXml;
using System.Threading;
using System.Diagnostics;
using System;
using System.Xml;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace Практика
{
    class Program
    {
        const string Name = "Отчет";
        const string Path = @"D:\" + Name + ".xlsx";
        const string ProcessName = Name + " - Excel";

        static void Main(string[] args)
        {
            CreateReport();
        }

        //Создание отчета
        public static void CreateReport()
        {
            Params Parameter = new Params();
            FileInfo File = new FileInfo(Path);
            Data SetData = new Data();

            if (File.Exists)
            {
                //Удаление имени Хабенского
                try
                {
                    File.Delete();
                }
                catch
                {
                    //Смерть процесса имени Хабенского
                    foreach (var KillProcess in Process.GetProcessesByName("excel"))
                    {
                        if (KillProcess.MainWindowTitle.Contains(ProcessName))
                        {
                            KillProcess.Kill();
                            break;
                        }
                    }
                    //Нужное для системы время на завершение процесса
                    Thread.Sleep(1000);
                    File.Delete();
                }
            }

            using (ExcelPackage Package = new ExcelPackage(File))
            {
                //Создание Книги и Страницы 1 Excel
                ExcelWorkbook WorkBook = Package.Workbook;
                ExcelWorksheet Worksheet1 = WorkBook.Worksheets.Add("Форма 1. Лист 1");

                //Рисование Страницы 1
                SetWorksheetSettings(Worksheet1);
                DrawWorksheet(Worksheet1, Parameter, 1);

                //Метод заполения данными из Xml
                SetData.DataFromXml(Parameter, Worksheet1, WorkBook);

                //Создание Страницы изменений Excel
                ExcelWorksheet Worksheet3 = WorkBook.Worksheets.Add("Лист изменений");

                //Рисование Страницы изменений
                SetWorksheetSettings(Worksheet3);
                DrawWorksheet(Worksheet3, Parameter, 3);

                //Сохранение имени Хабенского
                Package.Save();

                //Открытие имени того же Хабенского
                Process.Start(File.ToString());
            }
        }

        //Установка параметров печати и отображения страницы
        public static void SetWorksheetSettings(ExcelWorksheet Worksheet)
        {
            Worksheet.PrinterSettings.BottomMargin = Worksheet.PrinterSettings.LeftMargin = 0m;
            Worksheet.PrinterSettings.TopMargin = Worksheet.PrinterSettings.RightMargin = 0.21m;
            Worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            Worksheet.PrinterSettings.FitToHeight = 0;
            Worksheet.PrinterSettings.FitToWidth = 1;
            Worksheet.PrinterSettings.FitToPage = true;
            Worksheet.View.ShowGridLines = false;
            Worksheet.View.PageBreakView = true;     
        }

        //Рисовние таблицы по ее индексу
        public static void DrawWorksheet(ExcelWorksheet Worksheet, Params Parameter, int PageIndex)
        {
            if (PageIndex == 1)
            {
                CreaateColumns(Worksheet, Parameter, Parameter.P1ColumnWidth);
                CreateRows(Worksheet, Parameter, Parameter.P1RowHeight);
                MergeCells(Worksheet, Parameter.P1MergeIndexY1, Parameter.P1MergeIndexY2, Parameter.P1MergeIndexX1, Parameter.P1MergeIndexX2);
                CreateBorders(Worksheet, PageIndex, Parameter.P1BorderIndex);
                FillText(Worksheet, PageIndex, Parameter.P1FillingText, Parameter.P1TextAlignment, Parameter.P1TextRotation, Parameter.P3TextWrap);
            }
            else if (PageIndex == 2)
            {
                CreaateColumns(Worksheet, Parameter, Parameter.P2ColumnWidth);
                CreateRows(Worksheet, Parameter, Parameter.P2RowHeight);
                MergeCells(Worksheet, Parameter.P2MergeIndexY1, Parameter.P2MergeIndexY2, Parameter.P2MergeIndexX1, Parameter.P2MergeIndexX2);
                CreateBorders(Worksheet, PageIndex, Parameter.P2BorderIndex);
                FillText(Worksheet, PageIndex, Parameter.P2FillingText, Parameter.P2TextAlignment, Parameter.P2TextRotation, Parameter.P3TextWrap);
            }
            else if (PageIndex == 3)
            {
                CreaateColumns(Worksheet, Parameter, Parameter.P3ColumnWidth);
                CreateRows(Worksheet, Parameter, Parameter.P3RowHeight);
                MergeCells(Worksheet, Parameter.P3MergeIndexY1, Parameter.P3MergeIndexY2, Parameter.P3MergeIndexX1, Parameter.P3MergeIndexX2);
                CreateBorders(Worksheet, PageIndex, Parameter.P3BorderIndex);
                FillText(Worksheet, PageIndex, Parameter.P3FillingText, Parameter.P3TextAlignment, Parameter.P3TextRotation, Parameter.P3TextWrap);
            }
        }


        //Создание столбцов по параметрам
        public static void CreaateColumns(ExcelWorksheet Worksheet, Params Parameter, byte[] ColumnWidth)
        {
            int ColumnIndex = 0;
            for (int i = 0; i < ColumnWidth.Length; i++)
            {
                Worksheet.Column(++ColumnIndex).Width = Parameter.ToWidth(ColumnWidth[i]);
            }
        }

        //Создание строк по параметрам
        public static void CreateRows(ExcelWorksheet Worksheet, Params Parameter, byte[] RowHeight)
        {
            int RowIndex = 0;
            for (int i = 0; i < RowHeight.Length; i++)
            {
                Worksheet.Row(++RowIndex).Height = Parameter.ToHeight(RowHeight[i]);
            }
        }

        //Объединение ячеек по заданным индексам
        public static void MergeCells(ExcelWorksheet Worksheet, byte[] MergeIndexY1, byte[] MergeIndexY2, byte[] MergeIndexX1, byte[] MergeIndexX2)
        {
            for (int i = 0; i < MergeIndexY1.Length; i++)
            {
                Worksheet.Cells[MergeIndexY1[i], MergeIndexX1[i], MergeIndexY2[i], MergeIndexX2[i]].Merge = true;
            }
        }

        //Рисование границ по заданным индексам
        public static void CreateBorders(ExcelWorksheet Worksheet, int PageIndex, byte[] BorderIndex)
        {
            if (PageIndex == 1)
            {
                for (int i = 0; i < BorderIndex.Length; i += 4)
                {
                    using (ExcelRange Range = Worksheet.Cells[BorderIndex[i], BorderIndex[i + 1], BorderIndex[i + 2], BorderIndex[i + 3]])
                    {
                        Range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        Range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        Range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        Range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    }
                }

                Worksheet.Cells[36, 21].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                Worksheet.Cells[45, 21].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                Worksheet.Cells[45, 21].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                Worksheet.Cells[41, 12].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                Worksheet.Cells[36, 4, 37, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                Worksheet.Cells[43, 21, 44, 21].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                Worksheet.Cells[45, 4, 45, 20].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            }
            else
            {
                for (int i = 0; i < BorderIndex.Length; i += 4)
                {
                    using (ExcelRange Range = Worksheet.Cells[BorderIndex[i], BorderIndex[i + 1], BorderIndex[i + 2], BorderIndex[i + 3]])
                    {
                        Range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        Range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        Range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        Range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    }
                }
            }
        }

        //Заполнение ячеек стандартным текстом
        public static void FillText(ExcelWorksheet Worksheet, int PageIndex, string[,] FillingText, string[,] TextAlignment, string[] TextRotation, byte[] TextWrap)
        {
            if (PageIndex != 3)
            {
                Worksheet.Cells[TextRotation[TextRotation.Length - 1]].Style.WrapText = true;
            }
            else
            {
                Worksheet.Cells[TextWrap[0], TextWrap[1], TextWrap[2], TextWrap[3]].Style.WrapText = true;
            }

            Worksheet.Cells[1, 1, Worksheet.Dimension.End.Row + 1, Worksheet.Dimension.End.Column + 1].Style.Font.Italic = true;
            Worksheet.Cells[1, 1, Worksheet.Dimension.End.Row + 1, Worksheet.Dimension.End.Column + 1].Style.Font.Name = "GOST type B";
            Worksheet.Cells[1, 1, Worksheet.Dimension.End.Row + 1, Worksheet.Dimension.End.Column + 1].Style.Font.Size = 10;
            SetInsertedTextSettings(Worksheet, PageIndex, TextWrap);

            for (int i = 0; i < (FillingText.Length / 2); i++)
            {
                Worksheet.Cells[FillingText[0, i]].RichText.Add(FillingText[1, i]);
                Worksheet.Cells[FillingText[0, i]].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells[FillingText[0, i]].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }

            for (int i = 0; i < (TextAlignment.Length / 2); i++)
            {
                if (TextAlignment[1, i] == "Right")
                {
                    Worksheet.Cells[TextAlignment[0, i]].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
                else if (TextAlignment[1, i] == "Left")
                {
                    Worksheet.Cells[TextAlignment[0, i]].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                }
            }

            if (PageIndex != 3)
            {
                for (int i = 0; i < TextRotation.Length - 1; i++)
                {
                    Worksheet.Cells[TextRotation[i]].Style.TextRotation = 90;
                }
            }
            else
            {
                for (int i = 0; i < TextRotation.Length; i++)
                {
                    Worksheet.Cells[TextRotation[i]].Style.TextRotation = 90;
                }
            }
        }

        //Установка параметров для вставляемого текста
        public static void SetInsertedTextSettings(ExcelWorksheet Worksheet, int PageIndex, byte[] TextWrap)
        {
            if (PageIndex == 1)
            {
                Worksheet.Cells[3, 7, 35, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells[3, 7, 35, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["B18"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["B18"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["L38"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["L38"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["L41"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["L41"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["U42"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["C2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["U42"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["C2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells[TextWrap[4], TextWrap[5], TextWrap[6], TextWrap[7]].Style.WrapText = true;
                Worksheet.Cells[3, 4, 35, 21].Style.Font.Size = 12;
                Worksheet.Cells[2, 5, 2, 21].Style.Font.Size = 14;
                Worksheet.Cells["L38"].Style.Font.Size = 16;
                Worksheet.Cells["L41"].Style.Font.Size = 14;
            }
            else if (PageIndex == 2)
            {
                Worksheet.Cells[3, 7, 40, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells[3, 7, 40, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["B16"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["B16"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["L42"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["L42"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["P44"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["P44"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells[TextWrap[8], TextWrap[9], TextWrap[10], TextWrap[11]].Style.WrapText = true;
                Worksheet.Cells["L42"].Formula = "= 'Форма 1. Лист 1'!L38";
                Worksheet.Cells["J46"].Formula = "= 'Форма 1. Лист 1'!J46";
                Worksheet.Cells[3, 4, 40, 16].Style.Font.Size = 12;
                Worksheet.Cells[2, 5, 2, 16].Style.Font.Size = 14;
                Worksheet.Cells["L42"].Style.Font.Size = 16;
            }
            else if (PageIndex == 3)
            {
                Worksheet.Cells["S43"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                Worksheet.Cells["S43"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                Worksheet.Cells["S43"].Formula = "='Форма 1. Лист 1'!U42";
                Worksheet.Cells["I45"].Formula = "='Форма 1. Лист 1'!J46";
                Worksheet.Cells["D2"].Style.Font.Size = 14;
            }
        }
    }
}