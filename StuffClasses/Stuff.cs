using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VKSMM.ModelClasses;//Файл с классами моделей данных
using Excel = Microsoft.Office.Interop.Excel;

namespace VKSMM.StuffClasses
{
    class Stuff
    {

        //Декодер действия
        public static int ActionDecoder(string CodeAct)
        {
            if (CodeAct == "Блокировать") return 5;
            if (CodeAct == "Заменять") return 3;
            if (CodeAct == "Дописывать") return 4;
            if (CodeAct == "Пропускать") return 2;
            if (CodeAct == "Удалять") return 1;
            return -1;
        }

        //Кодер действия
        public static string ActionCoder(int CodeAct)
        {
            if (CodeAct == 5) return "Блокировать";
            if (CodeAct == 3) return "Заменять";
            if (CodeAct == 4) return "Дописывать";
            if (CodeAct == 2) return "Пропускать";
            if (CodeAct == 1) return "Удалять";
            return "Пусто";
        }

        public static List<string> ConvertMassToList(string[] Lines)
        {
            List<string> OutLine = new List<string>();

            foreach (string S in Lines)
            {
                OutLine.Add(S);
            }

            return OutLine;
        }

        public static int[] ConvertMassToInt(string[] Lines)
        {
            int[] OutLine = new int[Lines.Length];

            for (int i = 0; i < Lines.Length; i++)
            {
                OutLine[i] = Convert.ToInt32(Lines[i]);
            }

            return OutLine;
        }

        public static void UpdatePostavshikov(MainForm mainForm)
        {
            //mainForm.comboBox5.Items.Clear();

            //mainForm.comboBox5.Items.Add("ВСЕ");
            //mainForm.comboBox5.Text = "ВСЕ";

            foreach (Product P in mainForm.ProductListSource)
            {
                bool gift = true;

                string L = P.IDURL.Substring(P.IDURL.IndexOf("/id") + 1);

                L = L.Substring(0, L.IndexOf("?"));


                //foreach (object c in mainForm.comboBox5.Items)
                //{
                //    if (c.ToString() == L)
                //    {
                //        gift = false;
                //    }
                //}

                //if (gift)
                //{
                //    mainForm.comboBox5.Items.Add(L);
                //}
            }
        }

        // Экспорт данных из Excel-файла (не более 5 столбцов и любое количество строк <= 50.
        public static int ExportExcel(string FilePath, MainForm mainForm)
        {

            //int countf = 888000;

            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(FilePath);//ofd.FileName_PhotoPath



            Excel.Worksheet ObjWorkSheet_1 = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
            var lastCell_1 = ObjWorkSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                    // размеры базы
            int lastColumn_1 = (int)lastCell_1.Column;
            int lastRow_1 = (int)lastCell_1.Row;
            // Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 


            //MessageBox.Show("Данные из XLS получены!");


            DateTime dateTime = DateTime.Now;




            for (int i = 1; i < lastRow_1; i++) // по всем строкам
            {
                Action S1 = () => mainForm.label3.Text = "Загружено " + i + " из " + lastRow_1;
                mainForm.label3.Invoke(S1);



                Product prod = new Product();

                prod.IDURL = ObjWorkSheet_1.Cells[i + 1, 4].Text.ToString();


                string lll = ObjWorkSheet_1.Cells[i + 1, 1].Text.ToString();
                // lll = "g:\\Job\\Education\\VKSMM\\ТЕСТ\\ФОТО\\" + lll.Substring(lll.IndexOf("\\") + 1);
                lll = mainForm._PhotoPath + "\\" + lll.Substring(lll.IndexOf("\\") + 1);//fbd.SelectedPath

                prod.FilePath.Add(lll);


                prod.URLPhoto.Add(ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString());
                prod.datePost = Convert.ToDateTime(ObjWorkSheet_1.Cells[i + 1, 3].Text.ToString());

                try
                {
                    prod.prise[0] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 5].Text.ToString());
                    prod.prise[1] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 6].Text.ToString());
                    prod.prise[2] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 7].Text.ToString());
                    prod.prise[3] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 8].Text.ToString());
                    prod.prise[4] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 9].Text.ToString());
                }
                catch { }


                try
                {
                    string www = ObjWorkSheet_1.Cells[i + 1, 11].Text.ToString();
                    while (www.IndexOf("\n") >= 0)
                    {
                        prod.sellerText.Add(www.Substring(0, www.IndexOf("\n")));
                        www = www.Substring(www.IndexOf("\n") + 1);
                    }


                    if ((www.Length > 30) && (www.IndexOf(".") > 0))
                    {
                        string wwwb = www;
                        while (www.IndexOf(".") >= 0)
                        {
                            prod.sellerText.Add(www.Substring(0, www.IndexOf(".")));
                            www = www.Substring(www.IndexOf(".") + 1);
                        }
                    }

                    if (www.Length > 0)
                    {
                        prod.sellerText.Add(www);
                    }
                }
                catch { }


                prod.sellerTextCleen.Add(ObjWorkSheet_1.Cells[i + 1, 12].Text.ToString());

                bool get = false;
                int iii = 0;
                foreach (Product p in mainForm.ProductListSourceBuffer)
                {
                    if (p.IDURL == prod.IDURL)
                    {
                        get = true;
                        break;
                    }
                    iii++;
                }
                if (get)
                {

                    try
                    {
                        FileInfo f = new FileInfo(prod.FilePath[0]);
                        if (f.Exists)
                        {
                            mainForm.ProductListSourceBuffer[iii].FilePath.Add(prod.FilePath[0]);
                            mainForm.ProductListSourceBuffer[iii].URLPhoto.Add(prod.URLPhoto[0]);
                        }
                        else
                        {

                            try
                            {
                                using (WebClient client = new WebClient())
                                {
                                    client.DownloadFile(new Uri(prod.URLPhoto[0]), prod.FilePath[0]);
                                }
                            }
                            catch { }


                            f.Refresh();
                            if (f.Exists)
                            {
                                mainForm.ProductListSourceBuffer[iii].FilePath.Add(prod.FilePath[0]);
                                mainForm.ProductListSourceBuffer[iii].URLPhoto.Add(prod.URLPhoto[0]);
                            }
                            else
                            {
                                mainForm.imageNoExist.Add(prod.FilePath[0]);
                            }
                        }
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.ToString());
                    }


                }
                else
                {
                    // ProductListSourceBuffer.Insert(0, prod);
                    mainForm.ProductListSourceBuffer.Add(prod);
                    //dataGridView5.Rows.Add(prod.IDURL);
                    //listBox3.Items.Add(prod.IDURL);
                }
            }



            Action S2 = () => mainForm.label3.Text = "Загружено " + lastRow_1 + " из " + lastRow_1 + " " + (DateTime.Now - dateTime).ToString();
            mainForm.label3.Invoke(S2);



            //Excel.Worksheet ObjWorkSheet_2 = (Excel.Worksheet)ObjWorkBook.Sheets[2]; //получить 1-й лист
            //var lastCell_2 = ObjWorkSheet_2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку




            //// размеры базы
            //int lastColumn_2 = (int)lastCell_2.Column;
            //int lastRow_2 = (int)lastCell_2.Row;
            //// Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 
            //for (int i = 1; i < lastRow_2; i++) // по всем строкам
            //{
            //    string[] LIN = new string[5];

            //    for (int j = 0; j < 5; j++) //по всем колонкам
            //    {
            //        LIN[j] = ObjWorkSheet_2.Cells[i + 1, j + 1].Text.ToString(); //считываем данные
            //    }

            //    dataGridView1.Rows.Add(LIN);
            //}




            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из Excel
            GC.Collect(); // убрать за собой
            return 0;
        }

        // Экспорт данных из Excel-файла (не более 5 столбцов и любое количество строк <= 50.
        public static int ExportProviderExcel(MainForm mainForm)
        {
            try
            {

                //int countf = 888000;

                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(mainForm._ProviderDir);//ofd.FileName_PhotoPath



                Excel.Worksheet ObjWorkSheet_1 = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
                var lastCell_1 = ObjWorkSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                        // размеры базы
                int lastColumn_1 = (int)lastCell_1.Column;
                int lastRow_1 = (int)lastCell_1.Row;
                // Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 


                //MessageBox.Show("Данные из XLS получены!");


                DateTime dateTime = DateTime.Now;




                for (int i = 1; i < lastRow_1; i++) // по всем строкам
                {


                    Action S2 = () => mainForm.label11.Text = "Поставщиков обработано " + i.ToString() + " из " + lastRow_1.ToString();
                    mainForm.label11.Invoke(S2);

                    try
                    {

                        CategoryOfProduct COFP = new CategoryOfProduct();
                        COFP.Name = ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString();

                        bool Reg = true;
                        int indexCat = -1;
                        int IOd = 0;
                        foreach (CategoryOfProduct C in mainForm.mainCategoryList)
                        {
                            if (C.Name == COFP.Name)
                            {
                                Reg = false;
                                indexCat = IOd;
                                break;
                            }
                            IOd++;
                        }



                        if (COFP.Name.Length > 0)
                        {
                            if (Reg)
                            {
                                //Создаем экземпляр ключа
                                Key kmc = new Key();
                                //Значение ключа
                                kmc.Value = ObjWorkSheet_1.Cells[i + 1, 3].Text.ToString();
                                //kmc.Value = KEYPL.ChildNodes[0].InnerText;
                                //Флаг активности ключа
                                kmc.IsActiv = true;

                                COFP.Keys.Add(kmc);


                                COFP.isProvider = false;





                                mainForm.listBox1.Items.Add(COFP.Name);

                                mainForm.mainCategoryList.Add(COFP);

                                string[] s = new string[2];

                                s[0] = COFP.Name;
                                s[1] = COFP.SubCategoty.Count.ToString();
                                mainForm.dataGridView7.Rows.Add(s);
                            }
                            else
                            {

                                //Создаем экземпляр ключа
                                Key kmc = new Key();
                                //Значение ключа
                                kmc.Value = ObjWorkSheet_1.Cells[i + 1, 3].Text.ToString();
                                //kmc.Value = KEYPL.ChildNodes[0].InnerText;
                                //Флаг активности ключа
                                kmc.IsActiv = true;

                                kmc.isProvider = false;

                                //COFP.Keys.Add(kmc);

                                mainForm.mainCategoryList[indexCat].Keys.Add(kmc);

                            }
                        }



                    }
                    catch { }


                }





                ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                ObjWorkExcel.Quit(); // выйти из Excel
                GC.Collect(); // убрать за собой
            }
            catch
            {

            }
            return 0;
        }



        public static void CreateExcel(MainForm mainForm, SaveFileDialog ofd)
        {


            DirectoryInfo directory = new DirectoryInfo(ofd.FileName.Substring(0, ofd.FileName.LastIndexOf("\\") + 1));






            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("У Вас не установлден Exel!!");
                return;
            }


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //Set Text-Wrap for all rows true//
            xlWorkSheet.Rows.WrapText = true;

            xlWorkSheet.Rows.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            xlWorkSheet.Rows.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

            xlWorkSheet.Columns["A:A"].ColumnWidth = 20.14;
            xlWorkSheet.Columns["B:B"].ColumnWidth = 29.43;
            xlWorkSheet.Columns["C:C"].ColumnWidth = 21.71;
            xlWorkSheet.Columns["D:D"].ColumnWidth = 24.00;

            xlWorkSheet.Columns["E:E"].ColumnWidth = 10.71;
            xlWorkSheet.Columns["F:F"].ColumnWidth = 10.71;
            xlWorkSheet.Columns["G:G"].ColumnWidth = 10.71;
            xlWorkSheet.Columns["H:H"].ColumnWidth = 10.71;
            xlWorkSheet.Columns["I:I"].ColumnWidth = 10.71;

            xlWorkSheet.Columns["J:J"].ColumnWidth = 21.29;
            xlWorkSheet.Columns["K:K"].ColumnWidth = 35.29;
            xlWorkSheet.Columns["L:L"].ColumnWidth = 35.29;

            xlWorkSheet.Columns["M:M"].ColumnWidth = 33.57;

            xlWorkSheet.Columns["N:N"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["O:O"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["P:P"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["Q:Q"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["R:R"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["S:S"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["T:T"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["U:U"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["V:V"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["W:W"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["X:X"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["Y:Y"].ColumnWidth = 8.43;
            xlWorkSheet.Columns["Z:Z"].ColumnWidth = 8.43;


            xlWorkSheet.Cells[1, 1] = "№";
            xlWorkSheet.Cells[1, 2] = "Фото";
            xlWorkSheet.Cells[1, 3] = "Дата поста";
            xlWorkSheet.Cells[1, 4] = "Ссылка на товар";
            xlWorkSheet.Cells[1, 5] = "Цена 1";
            xlWorkSheet.Cells[1, 6] = "Цена 2";
            xlWorkSheet.Cells[1, 7] = "Цена 3";
            xlWorkSheet.Cells[1, 8] = "Цена 4";
            xlWorkSheet.Cells[1, 9] = "Цена 5";
            xlWorkSheet.Cells[1, 10] = "Размеры";
            xlWorkSheet.Cells[1, 11] = "Описание поставщика";
            xlWorkSheet.Cells[1, 12] = "Чистое описание";
            xlWorkSheet.Cells[1, 13] = "Материал";
            xlWorkSheet.Cells[1, 14] = "ссылка 1";
            xlWorkSheet.Cells[1, 15] = "ссылка 2";
            xlWorkSheet.Cells[1, 16] = "ссылка 3";
            xlWorkSheet.Cells[1, 17] = "ссылка 4";
            xlWorkSheet.Cells[1, 18] = "ссылка 5";
            xlWorkSheet.Cells[1, 19] = "ссылка 6";
            xlWorkSheet.Cells[1, 20] = "ссылка 7";
            xlWorkSheet.Cells[1, 21] = "ссылка 8";
            xlWorkSheet.Cells[1, 22] = "ссылка 9";
            xlWorkSheet.Cells[1, 23] = "ссылка 10";
            xlWorkSheet.Cells[1, 24] = "ссылка 11";
            xlWorkSheet.Cells[1, 25] = "ссылка 12";
            xlWorkSheet.Cells[1, 26] = "ссылка 13";


            int i = 2;
            int t = 1;
            foreach (Product p in mainForm.ProductListForPosting)
            {

                Action S2 = () => mainForm.label15.Text = "Постов обработано " + t.ToString() + " из " + mainForm.ProductListForPosting.Count.ToString();
                mainForm.label15.Invoke(S2);
                t++;

                string Opesanie1 = "";
                foreach (string ss in p.sellerText)
                {
                    Opesanie1 = Opesanie1 + ss + "\r\n";
                }

                string Opesanie2 = "";
                foreach (string ss in p.sellerTextCleen)
                {
                    Opesanie2 = Opesanie2 + ss + "\r\n";
                }


                int j = 0;
                foreach (string s in p.FilePath)
                {

                    string OutJPGPath = s;

                    try
                    {
                        Directory.CreateDirectory(directory.FullName + p.CategoryOfProductName + "\\" + p.SubCategoryOfProductName + "\\");

                        File.Copy(s, directory.FullName + p.CategoryOfProductName + "\\" + p.SubCategoryOfProductName + "\\" + s.Substring(s.LastIndexOf("\\") + 1), true);

                        OutJPGPath = directory.FullName + p.CategoryOfProductName + "\\" + p.SubCategoryOfProductName + "\\" + s.Substring(s.LastIndexOf("\\") + 1);

                    }
                    catch { }






                    xlWorkSheet.Cells[i, 1] = OutJPGPath;//s
                    xlWorkSheet.Cells[i, 2] = p.URLPhoto[j];
                    xlWorkSheet.Cells[i, 3] = p.datePost.ToString("dd/MM/yy hh:mm");
                    xlWorkSheet.Cells[i, 4] = p.IDURL;

                    xlWorkSheet.Cells[i, 5] = p.prise[0];
                    xlWorkSheet.Cells[i, 6] = p.prise[1];
                    xlWorkSheet.Cells[i, 7] = p.prise[2];
                    xlWorkSheet.Cells[i, 8] = p.prise[3];
                    xlWorkSheet.Cells[i, 9] = p.Prises;//p.prise[4];

                    xlWorkSheet.Cells[i, 10] = p.Sizes;
                    xlWorkSheet.Cells[i, 11] = Opesanie1;
                    xlWorkSheet.Cells[i, 12] = Opesanie2;
                    xlWorkSheet.Cells[i, 13] = p.Materials;
                    xlWorkSheet.Cells[i, 14] = "";
                    xlWorkSheet.Cells[i, 15] = "";
                    xlWorkSheet.Cells[i, 16] = "";
                    xlWorkSheet.Cells[i, 17] = "";
                    xlWorkSheet.Cells[i, 18] = "";
                    xlWorkSheet.Cells[i, 19] = "";
                    xlWorkSheet.Cells[i, 20] = "";
                    xlWorkSheet.Cells[i, 21] = "";
                    xlWorkSheet.Cells[i, 22] = "";
                    xlWorkSheet.Cells[i, 23] = "";
                    xlWorkSheet.Cells[i, 24] = "";
                    xlWorkSheet.Cells[i, 25] = "";
                    xlWorkSheet.Cells[i, 26] = "";

                    j++;
                    i++;
                }
            }



            //Захватываем диапазон ячеек
            //  Excel.Range range1 = xlWorkSheet.get_Range((Excel.Range)(xlWorkSheet.Cells[1, 1]), (Excel.Range)(xlWorkSheet.Cells[i, 1]));

            //  range1.ColumnWidth = 145;

            // (xlWorkSheet.Cells[1, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            // (xlWorkSheet.Cells[1, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            // (xlWorkSheet.Cells[1, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;





            //Захватываем другой диапазон ячеек
            Excel.Range range2 = xlWorkSheet.get_Range((Excel.Range)(xlWorkSheet.Cells[1, 1]), (Excel.Range)(xlWorkSheet.Cells[1, 26]));

            range2.Cells.Font.Name = "Calibri";
            range2.Cells.Font.Size = 14;
            range2.Cells.Font.Bold = true;

            //Задаем цвет этого диапазона. Необходимо подключить System.Drawing
            range2.Cells.Font.Color = ColorTranslator.ToOle(Color.White);
            //Фоновый цвет
            range2.Interior.Color = ColorTranslator.ToOle(Color.DarkGray);


            //Захватываем диапазон ячеек
            Excel.Range rangeAll = xlWorkSheet.get_Range((Excel.Range)(xlWorkSheet.Cells[1, 1]), (Excel.Range)(xlWorkSheet.Cells[i, 26]));
            //rangeAll.WrapText = Excel.wr
            rangeAll.EntireRow.AutoFit();


            //xlWorkSheet.Cells[1, 2] = "Name";
            //xlWorkSheet.Cells[2, 1] = "1";
            //xlWorkSheet.Cells[2, 2] = "One";
            //xlWorkSheet.Cells[3, 1] = "2";
            //xlWorkSheet.Cells[3, 2] = "Two";

            //	Фото Дата поста Ссылка на товар Цена 1  Цена 2  Цена 3  Цена 4  Цена 5  Размеры Описание поставщика Чистое описание Материал    ссылка 1    ссылка 2    ссылка 3    ссылка 4    ссылка 5    ссылка 6    ссылка 7    ссылка 8    ссылка 9    ссылка 10   ссылка 11   ссылка 12   ссылка 13



            //Here saving the file in xlsx + ".xlsx"
            xlWorkBook.SaveAs(ofd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Файл сохранен!");


            mainForm.ProductListForPosting.Clear();
            mainForm.treeView1.Nodes.Clear();
            mainForm.listView2.Items.Clear();

        }






    }
}
