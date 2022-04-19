using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using VKSMM.StuffClasses;//Файл с классами вспомогательных методов
using VKSMM.ModelClasses;//Файл с классами моделей данных

namespace VKSMM.ThredsCode
{
    /// <summary>
    /// Класс с телами основных потоков программы
    /// </summary>
    class Core
    {
        /// <summary>
        /// Тело процесса конвертации XML файлов с товаром
        /// </summary>
        public static void Thread_Dir_Processing_Code(object mainWindow)
        {

            MainForm mainForm = (MainForm)mainWindow;

            //Action S1 = () => button1.Enabled = false;
            //button1.Invoke(S1);

            DirectoryInfo d = new DirectoryInfo(mainForm._InputPath);

            if (!d.Exists)
            {
                MessageBox.Show("Директория " + mainForm._InputPath + " не существует!");
            }
            else
            {

                int cf = d.GetFiles().Length;
                int pf = 0;

                mainForm.imageNoExist.Clear();


                foreach (FileInfo f in d.GetFiles())
                {
                    try
                    {
                        Action S2 = () => mainForm.XMLLoadLabel.Text = "Файлов обработано " + pf.ToString() + " из " + cf.ToString();
                        mainForm.XMLLoadLabel.Invoke(S2);



                        Stuff.ExportExcel(f.FullName, mainForm);

                        //ExportExcel(f.FullName);



                        f.Delete();
                        pf++;
                    }
                    catch
                    {
                        //Произошла ошибка. Сообщаем оператору.  
                        MessageBox.Show("При конвертации файла: "+ f.FullName+" произошла ошибка! Проверьте файл!");
                    }
                }


                foreach (Product p in mainForm.ProductListSourceBuffer)
                {
                    mainForm.productListSource.Add(p);
                }


                mainForm.ProductListSourceBuffer.Clear();


                string sL = "При загрузке отсутствуют следующие изображения: \r\n";
                foreach (string L in mainForm.imageNoExist)
                {
                    sL = sL + L + "\r\n";
                }

                MessageBox.Show(sL);

                Action S3 = () => mainForm.productUnProcessedListBox.Items.Clear(); 
                mainForm.productUnProcessedListBox.Invoke(S3);

                //mainForm.productUnProcessedListBox.Items.Clear();

                int i = 1;

                foreach (Product P in mainForm.productListSource)
                {
                    Action S4 = () => mainForm.productUnProcessedListBox.Items.Add(i);
                    mainForm.productUnProcessedListBox.Invoke(S4);

                    
                    i++;
                }

                //Stuff.UpdatePostavshikov(mainForm);

                //Action S3 = () => button1.Enabled = true;
                //button1.Invoke(S3);
            }

        }

        /// <summary>
        /// Тело процесса сохранения товаров для поста в XML файлов с товаром
        /// </summary>
        public static void Thread_Create_Excel_Code(object outputForm)
        {
            VKSMM.OutputForm outForm  = (VKSMM.OutputForm)outputForm;


            DirectoryInfo directory = new DirectoryInfo(outForm.outputFilePath.Substring(0, outForm.outputFilePath.LastIndexOf("\\") + 1));

            Action S1 = () => outForm.labelOutputFormResult.Text = "ЗАГРУЖАЕТСЯ XLS";
            outForm.labelOutputFormResult.Invoke(S1);

            

        

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

            xlWorkSheet.Rows.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xlWorkSheet.Rows.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

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



            Action S2 = () => outForm.labelOutputFormResult.Text = "ВЫГРУЗКА ТОВАРА";
            outForm.labelOutputFormResult.Invoke(S2);


            //int countProduct = outForm.mainForm.ProductListForPosting.Count;
            
            Action S3 = () => outForm.progressBarOutputForm.Maximum = outForm.mainForm.ProductListForPosting.Count;
            outForm.progressBarOutputForm.Invoke(S3);


            int i = 2;
            int t = 1;
            foreach (Product p in outForm.mainForm.ProductListForPosting)
            {

                Action S4 = () => outForm.progressBarOutputForm.Value = t;
                outForm.progressBarOutputForm.Invoke(S4);

                //Action S2 = () => outForm.mainForm.label15.Text = "Постов обработано " + t.ToString() + " из " + outForm.mainForm.ProductListForPosting.Count.ToString();
                //outForm.mainForm.label15.Invoke(S2);

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
                        if (outForm.upLoadImage)
                        {
                            Directory.CreateDirectory(directory.FullName + p.CategoryOfProductName + "\\" + p.SubCategoryOfProductName + "\\");

                            File.Copy(s, directory.FullName + p.CategoryOfProductName + "\\" + p.SubCategoryOfProductName + "\\" + s.Substring(s.LastIndexOf("\\") + 1), true);
                        }

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


            Action S5 = () => outForm.labelOutputFormResult.Text = "СОХРАНЕНИЕ XML";
            outForm.labelOutputFormResult.Invoke(S5);



            //Захватываем другой диапазон ячеек
            Microsoft.Office.Interop.Excel.Range range2 = xlWorkSheet.get_Range((Microsoft.Office.Interop.Excel.Range)(xlWorkSheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(xlWorkSheet.Cells[1, 26]));

            range2.Cells.Font.Name = "Calibri";
            range2.Cells.Font.Size = 14;
            range2.Cells.Font.Bold = true;

            //Задаем цвет этого диапазона. Необходимо подключить System.Drawing
            range2.Cells.Font.Color = ColorTranslator.ToOle(Color.White);
            //Фоновый цвет
            range2.Interior.Color = ColorTranslator.ToOle(Color.DarkGray);


            //Захватываем диапазон ячеек
            Microsoft.Office.Interop.Excel.Range rangeAll = xlWorkSheet.get_Range((Microsoft.Office.Interop.Excel.Range)(xlWorkSheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(xlWorkSheet.Cells[i, 26]));
            //rangeAll.WrapText = Excel.wr
            rangeAll.EntireRow.AutoFit();


            //xlWorkSheet.Cells[1, 2] = "Name";
            //xlWorkSheet.Cells[2, 1] = "1";
            //xlWorkSheet.Cells[2, 2] = "One";
            //xlWorkSheet.Cells[3, 1] = "2";
            //xlWorkSheet.Cells[3, 2] = "Two";

            //	Фото Дата поста Ссылка на товар Цена 1  Цена 2  Цена 3  Цена 4  Цена 5  Размеры Описание поставщика Чистое описание Материал    ссылка 1    ссылка 2    ссылка 3    ссылка 4    ссылка 5    ссылка 6    ссылка 7    ссылка 8    ссылка 9    ссылка 10   ссылка 11   ссылка 12   ссылка 13



            //Here saving the file in xlsx + ".xlsx"
            xlWorkBook.SaveAs(outForm.outputFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            //MessageBox.Show("Файл сохранен!");
            Action S6 = () => outForm.labelOutputFormResult.Text = "ЗАВЕРШЕНО!";
            outForm.labelOutputFormResult.Invoke(S6);


            outForm.mainForm.ProductListForPosting.Clear();
            
            Action S8 = () => outForm.mainForm.treeViewProductForPostBox.Nodes.Clear();
            outForm.mainForm.treeViewProductForPostBox.Invoke(S8);
            Action S9 = () => outForm.mainForm.listViewPostBox.Items.Clear();
            outForm.mainForm.listViewPostBox.Invoke(S9);


            Action S10 = () => outForm.Close();
            outForm.Invoke(S10);

        }

        /// <summary>
        /// Тело процесса конвертации XML файла провайдеров
        /// </summary>
        public static void Thread_Provider_Excel_Code(object parametr)
        {
            VKSMM.LoadForm loadForm = (VKSMM.LoadForm)parametr;

            try
            {

                //int countf = 888000;

                Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(loadForm.mainForm._ProviderDir);//ofd.FileName_PhotoPath



                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet_1 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
                var lastCell_1 = ObjWorkSheet_1.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
                                                                                                        // размеры базы
                int lastColumn_1 = (int)lastCell_1.Column;
                int lastRow_1 = (int)lastCell_1.Row;
                // Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 


                //MessageBox.Show("Данные из XLS получены!");


                DateTime dateTime = DateTime.Now;


                Action S1 = () => loadForm.progressBarLoadForm.Maximum = lastRow_1;
                loadForm.progressBarLoadForm.Invoke(S1);


                for (int i = 1; i < lastRow_1; i++) // по всем строкам
                {

                    string[] line = new string[4];
                    line[0] = ObjWorkSheet_1.Cells[i + 1, 1].Text.ToString();
                    line[1] = ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString();
                    line[2] = ObjWorkSheet_1.Cells[i + 1, 3].Text.ToString();
                    line[3] = ObjWorkSheet_1.Cells[i + 1, 4].Text.ToString();
                    loadForm.mainForm.providerDataGrid.Rows.Add(line);




                    Action S2 = () => loadForm.labelLoadForm.Text = "Поставщиков обработано " + i.ToString() + " из " + lastRow_1.ToString();
                    loadForm.labelLoadForm.Invoke(S2);

                    Action S3 = () => loadForm.progressBarLoadForm.Value = i;
                    loadForm.progressBarLoadForm.Invoke(S3);



                    try
                    {

                        CategoryOfProduct COFP = new CategoryOfProduct();
                        COFP.Name = ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString();

                        bool Reg = true;
                        int indexCat = -1;
                        int IOd = 0;
                        foreach (CategoryOfProduct C in loadForm.mainForm.providerCategoryList)
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





                                //mainForm.catListBox.Items.Add(COFP.Name);

                                loadForm.mainForm.providerCategoryList.Add(COFP);

                                //string[] s = new string[2];

                                //s[0] = COFP.Name;
                                //s[1] = COFP.SubCategoty.Count.ToString();
                                //mainForm.dataGridView7.Rows.Add(s);
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

                                loadForm.mainForm.providerCategoryList[indexCat].Keys.Add(kmc);

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
            //return 0;

            Action S4 = () => loadForm.Close();
            loadForm.Invoke(S4);

        }

    }
}
