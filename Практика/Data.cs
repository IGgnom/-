using System;
using System.Xml;
using System.Linq;
using OfficeOpenXml;
using System.Collections;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace Практика
{
    class Data
    {
        //Список для сортировки всех данных
        public static ArrayList DataList = new ArrayList();

        //Главный метод. Заполняет данными из заголовка и вызывает заполнение основными данными
        public void DataFromXml(Params Parameter, ExcelWorksheet Worksheet, ExcelWorkbook WorkBook)
        {
            XmlDocument Document = new XmlDocument();
            Document.Load(@"D:\VSProjects\Практика\SpecReport.xml");
            XmlElement Root = Document.DocumentElement;
            Parameter.SecondWorksheet = false;
            Parameter.ToNextPage = false;
            Parameter.MoveIndex = 3;
            Parameter.Counter = 0;
            Parameter.Length = 33;
            Parameter.Page = 2;
            int HeadIndex = 0;

            foreach (XmlNode Node in Root)
            {
                foreach (XmlNode ChildNode1 in Node.ChildNodes)
                {
                    foreach (XmlNode ChildNode2 in ChildNode1.ChildNodes)
                    {
                        if (ChildNode1.Name != "val")
                        {
                            XmlNode Attributes1 = ChildNode1.Attributes.GetNamedItem("Name");
                            Parameter.Razdel = Attributes1.Value;
                        }
                        else
                        {
                            XmlNode Attributes = ChildNode1.Attributes.GetNamedItem("title");
                            if (ChildNode1.InnerText != "" && Attributes.Value != "Logo")
                            {
                                if (Attributes.Value == "Razrabotal" || Attributes.Value == "Proveril" || Attributes.Value == "Normo_Kontr" || Attributes.Value == "Utv")
                                {
                                    string[] NodeText = ChildNode1.InnerText.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                    Worksheet.Cells[Parameter.Header[HeadIndex++]].RichText.Add(NodeText[0]);
                                }
                                else if (Attributes.Value == "CRC")
                                {
                                    Worksheet.Cells["J46"].RichText.Add(ChildNode1.InnerText);
                                }
                                else
                                {
                                    string[] NodeText = new string[] { ChildNode1.InnerText };
                                    Worksheet.Cells[Parameter.Header[HeadIndex++]].RichText.Add(NodeText[0]);
                                }
                            }
                        }
                        Parameter.SetNullVariables();
                        foreach (XmlNode ChildNode3 in ChildNode2.ChildNodes)
                        {
                            XmlNode Attributes2 = ChildNode3.Attributes.GetNamedItem("title");
                            Parameter.SetData(ChildNode3, Attributes2);
                        }
                        InsertData(Parameter.Razdel, Parameter.Format, Parameter.Obozn, Parameter.Naimen0, Parameter.Order, Parameter.Prim, Parameter.Pos, Parameter.Naimen1, Parameter.Naimen2, Parameter.Naimen3, Parameter.Kol, Parameter.EdIzm);
                    }
                }
            }
            SortAndInsert(WorkBook, Worksheet, Parameter); 
        }

        //Метод сортировки списка с последующей вставкой из него
        public static void SortAndInsert(ExcelWorkbook WorkBook, ExcelWorksheet Worksheet, Params Parameter)
        { 
            for (int i = 0; i < DataList.ToArray().Length; i++)
                for (int j = 0; j < DataList.ToArray().Length - 1 - i; j++)
                    if (Convert.ToInt32(DataList[j].ToString().Substring(0, DataList[j].ToString().IndexOf('$'))) > Convert.ToInt32(DataList[j + 1].ToString().Substring(0, DataList[j + 1].ToString().IndexOf('$'))))
                    {
                        string TempData = DataList[j].ToString();
                        DataList[j] = DataList[j + 1];
                        DataList[j + 1] = TempData;
                    }

            for (int i = 0; i < DataList.ToArray().Length; i++)
            {
                if (Convert.ToInt32(DataList[i].ToString().Substring(0, DataList[i].ToString().IndexOf('$'))) == 0)
                {
                    DataList[i] = DataList[i].ToString().Substring(DataList[i].ToString().IndexOf('$'), DataList[i].ToString().Length - 1);
                }
            }

            for (int i = 0; i < DataList.ToArray().Length; i++)
            {
                if (i < DataList.ToArray().Length - 1)
                {
                    InsertRow(WorkBook, Worksheet, Parameter, DataList[i].ToString(), DataList[i + 1].ToString());
                }
                else
                {
                    InsertRow(WorkBook, Worksheet, Parameter, DataList[i].ToString());
                }
            }
            Worksheet.Cells["U42"].RichText.Add(Convert.ToString(Parameter.Page));
        }

        //Сам метод построчной вставки данных в Excel, при необходимости создающий новые Страницы
        public static void InsertRow(ExcelWorkbook WorkBook, ExcelWorksheet FirstWorksheet, Params Parameter, string InputString, string NextString = "$Документация$-$ТЕСТ.010101.010 П1М$Тестирование$0$")
        {
            string[] Rows = InputString.Split(new char[] { '$' });
            string[] RowsNext = NextString.Split(new char[] { '$' });
            string[] Split1 = Parameter.SplitRows(Rows[4], 26, 1);
            string[] Split2 = Parameter.SplitRows(Rows[6], 10, 2);
            string[] SplitNext1 = Parameter.SplitRows(RowsNext[4], 26, 1);
            string[] SplitNext2 = Parameter.SplitRows(RowsNext[6], 10, 2);
            int[] MovingCells1 = new int[] { 9, 18, 21, 27, 31 };
            int[] MovingCells2 = new int[] { 19, 25, 29, 33, 39 };
            int MoveIndex1 = Parameter.MoveIndex;
            int MoveIndex2 = Parameter.MoveIndex;
            int SplitNext = 0;
            bool NextRow = false;
            bool Move = false;
            if (Parameter.SplitNextOld + Parameter.MoveIndex < Parameter.Length)
            {
                if (SplitNext1.Length > SplitNext2.Length)
                    SplitNext = SplitNext1.Length;
                else
                    SplitNext = SplitNext2.Length;
            }
            else
            {
                SplitNext = Parameter.SplitNextOld;
            }
            if (Parameter.MoveIndex + SplitNext > Parameter.Length)
            {
                ExcelWorksheet Worksheet = WorkBook.Worksheets.Add("Форма 1а.Лист " + Convert.ToString(Parameter.Page++));
                Program.SetWorksheetSettings(Worksheet);
                Program.DrawWorksheet(Worksheet, Parameter, 2);
                Parameter.MoveIndex = MoveIndex1 = MoveIndex2 = 3;
                Parameter.Counter = Worksheet.Index;
                Parameter.ToNextPage = true;
                Parameter.Length = 39;
                if (Parameter.RazdelOld != Rows[1])
                {
                    if (MovingCells2.Contains(Parameter.MoveIndex))
                        Parameter.MoveIndex += 2;
                    else if (Parameter.MoveIndex == 18 || Parameter.MoveIndex == 24 || Parameter.MoveIndex == 28 || Parameter.MoveIndex == 32 || Parameter.MoveIndex == 38)
                    {
                        Parameter.MoveIndex += 2;
                        NextRow = true;
                    }
                    else
                        Parameter.MoveIndex++;
                    Parameter.RazdelOld = Rows[1];
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex++)].RichText.Add(Rows[1]);
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.Font.UnderLine = true;
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    if (MovingCells2.Contains(Parameter.MoveIndex))
                        Parameter.MoveIndex = MoveIndex1 = MoveIndex2 += 4;
                    else
                    {
                        MoveIndex1 = MoveIndex2 += 3;
                        Parameter.MoveIndex++;
                    }
                    if (NextRow == true)
                    {
                        NextRow = false;
                        MoveIndex1 = MoveIndex2 = ++Parameter.MoveIndex; ;
                    }
                }
                if (MovingCells2.Contains(Parameter.MoveIndex))
                    Parameter.MoveIndex++;
                Worksheet.Cells["P44"].RichText.Add(Convert.ToString(Parameter.Page - 1));
                Worksheet.Cells["D" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[2]);
                Worksheet.Cells["G" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[0]);
                Worksheet.Cells["I" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[3]);
                
                for (int i = 0; i < Split1.Length; i++)
                {
                    if (MovingCells2.Contains(MoveIndex1))
                        MoveIndex1++;
                    Worksheet.Cells["M" + Convert.ToString(MoveIndex1++)].RichText.Add(Split1[i]);
                }
                Worksheet.Cells["N" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[5]);
                for (int i = 0; i < Split2.Length; i++)
                {
                    if (MovingCells2.Contains(MoveIndex2))
                        MoveIndex2++;
                    Worksheet.Cells["O" + Convert.ToString(MoveIndex2++)].RichText.Add(Split2[i]);
                }
                if (MoveIndex2 > MoveIndex1)
                {
                    Parameter.MoveIndex = MoveIndex2;
                }
                else
                {
                    Parameter.MoveIndex = MoveIndex1;
                }
            }
            else if (Parameter.ToNextPage == true)
            {
                ExcelWorksheet Worksheet = WorkBook.Worksheets[Parameter.Counter];
                if (Parameter.RazdelOld != Rows[1])
                {
                    if (MovingCells2.Contains(Parameter.MoveIndex))
                        Parameter.MoveIndex += 2;
                    else if (Parameter.MoveIndex == 18 || Parameter.MoveIndex == 24 || Parameter.MoveIndex == 28 || Parameter.MoveIndex == 32 || Parameter.MoveIndex == 38)
                    {
                        Parameter.MoveIndex += 2;
                        NextRow = true;
                    }
                    else
                        Parameter.MoveIndex++;
                    Parameter.RazdelOld = Rows[1];
                    
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex++)].RichText.Add(Rows[1]);
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.Font.UnderLine = true;
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    Worksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    if (MovingCells2.Contains(Parameter.MoveIndex))
                        Parameter.MoveIndex = MoveIndex1 = MoveIndex2 += 4;
                    else
                    {
                        MoveIndex1 = MoveIndex2 += 3;
                        Parameter.MoveIndex++;
                    }
                    if (NextRow == true)
                    {
                        NextRow = false;
                        MoveIndex1 = MoveIndex2 = ++Parameter.MoveIndex; ;
                    }
                }
                if (MovingCells2.Contains(Parameter.MoveIndex))
                    Parameter.MoveIndex++;
                Worksheet.Cells["D" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[2]);
                Worksheet.Cells["G" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[0]);
                Worksheet.Cells["I" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[3]);
                for (int i = 0; i < Split1.Length; i++)
                {
                    if (MovingCells2.Contains(MoveIndex1))
                        MoveIndex1++;
                    else if (MoveIndex1 == 18)
                        Move = true;
                    if (Split1.Length > 1)
                        Move = false;
                    Worksheet.Cells["M" + Convert.ToString(MoveIndex1++)].RichText.Add(Split1[i]);
                }
                Worksheet.Cells["N" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[5]);
                for (int i = 0; i < Split2.Length; i++)
                {
                    if (MovingCells2.Contains(MoveIndex2))
                        MoveIndex2++;
                    Worksheet.Cells["O" + Convert.ToString(MoveIndex2++)].RichText.Add(Split2[i]);
                }
                if (Move == true)
                {
                    Parameter.MoveIndex = 20;
                    Move = false;
                }
                else
                {
                    if (MoveIndex2 > MoveIndex1)
                    {
                        Parameter.MoveIndex = MoveIndex2;
                    }
                    else
                    {
                        Parameter.MoveIndex = MoveIndex1;
                    }
                } 
            }
            else
            {
                if (Parameter.RazdelOld != Rows[1])
                {
                    if (MovingCells1.Contains(Parameter.MoveIndex))
                        Parameter.MoveIndex += 2;
                    else if (Parameter.MoveIndex == 8 || Parameter.MoveIndex == 17 || Parameter.MoveIndex == 20 || Parameter.MoveIndex == 26 || Parameter.MoveIndex == 31)
                    {
                        Parameter.MoveIndex += 2;
                        NextRow = true;
                    }
                    else
                        Parameter.MoveIndex++;
                    Parameter.RazdelOld = Rows[1];
                    FirstWorksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex++)].RichText.Add(Rows[1]);
                    FirstWorksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.Font.UnderLine = true;
                    FirstWorksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    FirstWorksheet.Cells["M" + Convert.ToString(Parameter.MoveIndex - 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    if (MovingCells1.Contains(Parameter.MoveIndex))
                        Parameter.MoveIndex = MoveIndex1 = MoveIndex2 += 4;
                    else
                    {
                        MoveIndex1 = MoveIndex2 += 3;
                        Parameter.MoveIndex++;
                        if (Parameter.MoveIndex == 12)
                        {
                            Parameter.MoveIndex--;
                        }      
                    }
                    if (NextRow == true)
                    {
                        NextRow = false;
                        MoveIndex1 = MoveIndex2 = ++Parameter.MoveIndex; ;
                    }
                }
                if (MovingCells1.Contains(Parameter.MoveIndex))
                Parameter.MoveIndex++;
                FirstWorksheet.Cells["D" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[2]);
                FirstWorksheet.Cells["G" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[0]);
                FirstWorksheet.Cells["I" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[3]);
                MoveIndex1 = MoveIndex2 = Parameter.MoveIndex;
                for (int i = 0; i < Split1.Length; i++)
                {
                    if (MovingCells1.Contains(MoveIndex1))
                        MoveIndex1++;
                    FirstWorksheet.Cells["M" + Convert.ToString(MoveIndex1++)].RichText.Add(Split1[i]);
                }
                FirstWorksheet.Cells["R" + Convert.ToString(Parameter.MoveIndex)].RichText.Add(Rows[5]);
                for (int i = 0; i < Split2.Length; i++)
                {
                    if (MovingCells1.Contains(MoveIndex2))
                        MoveIndex2++;
                    FirstWorksheet.Cells["T" + Convert.ToString(MoveIndex2++)].RichText.Add(Split2[i]);
                }
                if (MoveIndex2 > MoveIndex1)
                {
                    Parameter.MoveIndex = MoveIndex2;
                }
                else
                {
                    Parameter.MoveIndex = MoveIndex1;
                }
            }
            Parameter.SplitNextOld = SplitNext;
        }

        //Метод для добавления данных с разделителями в список 
        public static void InsertData(string Razdel, string Format, string Obozn, string Naimen0, string Order, string Prim, string Pos, string Naimen1, string Naimen2, string Naimen3, string Kol, string EdIzm)
        {
            if (Razdel != null)
            {
                string Naimen = "";
                string Kolvo = Kol;
                string Primech = Prim == "" ? EdIzm : Prim;
                if (Naimen0 == "" || Naimen0 == null)
                {
                    if (Naimen3 != null || Naimen3 != "")
                    {
                        Naimen = Naimen1 + " " + Naimen2 + " " + Naimen3;
                    }
                    else
                    {
                        Naimen = Naimen1 + " " + Naimen2;
                    }
                }
                else
                {
                    Naimen = Naimen0;
                }
                if (Pos == null)
                    Pos = Convert.ToString(0);
                string AddString = Pos + "$" + Razdel + "$" + Format + "$" + Obozn + "$" + Naimen + "$" + Kolvo + "$" + Primech;
                DataList.Add(AddString);
            }
        }
    }
}
