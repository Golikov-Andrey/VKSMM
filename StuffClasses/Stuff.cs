using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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

        //public static void UpdatePostavshikov(MainForm mainForm)
        //{
        //    //mainForm.comboBox5.Items.Clear();

        //    //mainForm.comboBox5.Items.Add("ВСЕ");
        //    //mainForm.comboBox5.Text = "ВСЕ";

        //    foreach (Product P in mainForm.productListSource)
        //    {
        //        bool gift = true;

        //        string L = P.IDURL.Substring(P.IDURL.IndexOf("/id") + 1);

        //        L = L.Substring(0, L.IndexOf("?"));


        //        //foreach (object c in mainForm.comboBox5.Items)
        //        //{
        //        //    if (c.ToString() == L)
        //        //    {
        //        //        gift = false;
        //        //    }
        //        //}

        //        //if (gift)
        //        //{
        //        //    mainForm.comboBox5.Items.Add(L);
        //        //}
        //    }
        //}

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



            //Проходим по всем строчкам XLS файла
            for (int i = 1; i < lastRow_1; i++) // по всем строкам
            {

                //if(i==196)
                //{
                //    MessageBox.Show("11");
                //}


                try
                {
                    //Сообщаем на форму, что обработана строчка
                    Action S1 = () => mainForm.providerLoadLabel.Text = "Загружено " + i + " из " + lastRow_1;
                    mainForm.providerLoadLabel.Invoke(S1);

                    //Создаем новый товар
                    Product prod = new Product();

                    //Ссылка на товар
                    prod.IDURL = ObjWorkSheet_1.Cells[i + 1, 4].Text.ToString();

                    //Получаем ссылку на URL
                    prod.URLPhoto.Add(ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString());


                    //Ссылка на фотографию товара
                    string lll = ObjWorkSheet_1.Cells[i + 1, 1].Text.ToString();
                    // lll = "g:\\Job\\Education\\VKSMM\\ТЕСТ\\ФОТО\\" + lll.Substring(lll.IndexOf("\\") + 1);
                    lll = mainForm._PhotoPath + "\\" + lll.Substring(lll.IndexOf("\\") + 1);//fbd.SelectedPath

                    //--------------------------------------------------------------------------------------------
                    //Проверяем существует ли фотография, докачиваем и проверяем на повторы
                    //--------------------------------------------------------------------------------------------
                    //Получаем ссылку на изображение
                    FileInfo fimage = new FileInfo(lll);
                    //Проверяем существует ли файл с картинкой
                    if (!fimage.Exists)
                    {
                        try
                        {
                            //Докачиваем изображение если его не существует
                            using (WebClient client = new WebClient())
                            {
                                client.DownloadFile(new Uri(prod.URLPhoto[0]), lll);// prod.FilePath[0]
                            }
                        }
                        catch { }

                        //Обновляем ссылку на файл с картинкой
                        fimage.Refresh();
                    }


                    bool imageRepeat = false;

                    //Если файл с картинкой существует проводим проверку на повтор
                    if (fimage.Exists)
                    {


                        //Получаем гистограмму фотографии
                        int[] histogramm = collectingHistogramm(new Bitmap(lll));


                        imageRepeat = repeatImageTest(histogramm, 0.05, mainForm);

                        if (imageRepeat)
                        {
                            mainForm.imageDoubleList.Add(lll);

                            Directory.CreateDirectory(mainForm._PhotoPath + "\\REPEAT_IMAGE\\");

                            fimage.CopyTo(mainForm._PhotoPath + "\\REPEAT_IMAGE\\" + fimage.Name, true);
                        }
                        else
                        {
                            prod.FilePath.Add(lll);
                            mainForm.imageHistogrammList.Add(histogramm);
                        }

                    }
                    //--------------------------------------------------------------------------------------------




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

                    try
                    {
                        prod.sellerTextCleen.Add(ObjWorkSheet_1.Cells[i + 1, 12].Text.ToString());
                    }
                    catch
                    {
                        prod.sellerTextCleen.Add(" ");
                    }

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
                        if (!imageRepeat)
                        {
                            try
                            {
                                mainForm.ProductListSourceBuffer[iii].FilePath.Add(prod.FilePath[0]);
                                mainForm.ProductListSourceBuffer[iii].URLPhoto.Add(prod.URLPhoto[0]);
                            }
                            catch
                            {
                                mainForm.ProductListSourceBuffer.Add(prod);
                            }
                        }
                    }
                    else
                    {
                        mainForm.ProductListSourceBuffer.Add(prod);
                    }
                }
                catch
                {
                    MessageBox.Show("11111");
                }
            }



            Action S2 = () => mainForm.providerLoadLabel.Text = "Загружено " + lastRow_1 + " из " + lastRow_1 + " " + (DateTime.Now - dateTime).ToString();
            mainForm.providerLoadLabel.Invoke(S2);



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






        public static int[] collectingHistogramm(Image image)
        {
            //Выходная гистограмма
            int[] outHistogramm = new int[256 + 256 + 256];


            // получаем битмап из изображения
            Bitmap bmp = new Bitmap(image);

            //// создаем массивы, в котором будут содержаться количества повторений для каждого из значений каналов.
            //// индекс соответствует значению канала
            //int[] R = new int[256];
            //int[] G = new int[256];
            //int[] B = new int[256];

            int i, j;
            Color color;

            // собираем статистику для изображения
            for (i = 0; i < bmp.Width; ++i)
                for (j = 0; j < bmp.Height; ++j)
                {
                    color = bmp.GetPixel(i, j);
                    ++outHistogramm[color.R];
                    ++outHistogramm[color.G + 256];
                    ++outHistogramm[color.B + 512];
                }

            return outHistogramm;
        }

        public static bool repeatImageTest(int[] histogramm, double parametr, MainForm mainForm)
        {
            bool result = false;


            foreach (int[] hist in mainForm.imageHistogrammList)
            {
                double sum = 0;

                for (int i = 0; i < 256 * 3; i++)
                {
                    sum = sum + (1 - (double)((double)Math.Min(hist[i], histogramm[i]) / Math.Max(hist[i], histogramm[i])));
                }

                sum = sum / (256 * 3);


                if (sum < parametr) 
                { 
                    result = true; 
                    break; 
                }

            }

            return result;
        }









        public static List<int> ProcessingProductsAll(MainForm mainForm)
        {
            int index = 0;
            //int indexTV = 0;
            int indexTV = mainForm.ProductListForPosting.Count;

            List<int> RemovedIndexes = new List<int>();

            foreach (Product P in mainForm.productListSource)
            {
                if (((P.IDURL.IndexOf(mainForm.comboBox5.Text) >= 0) || (mainForm.comboBox5.Text == "ВСЕ"))
                    && ((P.CategoryOfProductName == mainForm.comboBox4.Text) || (mainForm.comboBox4.Text == "ВСЕ")))
                {



                    P.sellerTextCleen.Clear();



                    bool isCat = false;
                    bool isSub = false;
                    bool isStop = true;


                    //Буфер с новым описанием
                    string stringpost = "";


                    bool goodsProcessed = true;

                    string Razmer = string.Empty;
                    bool RazmerFind = false;

                    List<string> akamulateRegexLog = new List<string>();


                    //Проходим по всем строчкам из описания
                    for (int u = 0; u < P.sellerText.Count; u++)//listBox2.SelectedIndex
                    {
                        string s = P.sellerText[u];

                        ////В строчке должны быть данные
                        //if (s.Length > 1)
                        //{
                        //    //Добавляем строчку описания в грид
                        //    dataGridView3.Rows.Add(s);
                        //}

                        //Ищем ключевые слова для начала сборки размеров
                        if ((s.ToLower().IndexOf("размер") >= 0) || (s.ToLower().IndexOf("разм.") >= 0) || (s.ToLower().IndexOf("рост") >= 0) || (s.ToLower().IndexOf("opct") >= 0))
                        {
                            //Ключи найдены поднимаем флаг сборки размеров
                            RazmerFind = true;

                            akamulateRegexLog.Add("# сборка размера начата");

                        }

                        //Действие при сборке размеров
                        if (RazmerFind)
                        {
                            //Аккамулируем строчки размера
                            Razmer = Razmer + " " + s;
                            akamulateRegexLog.Add("# сборка размера "+ Razmer);

                            //Если конец описание не достигнут то обрабатываем следующую строчку
                            if (u < P.sellerText.Count - 1)
                            {
                                //Если на следующей строчке есть ключ "рост" то блокируем сборку
                                if ((P.sellerText[u + 1].ToLower().IndexOf("рост") >= 0))
                                {
                                    s = Razmer + " " + P.sellerText[u + 1];

                                    u++;

                                    RazmerFind = false;
                                    akamulateRegexLog.Add("# сборка размера закончена - в следующей строчке есть |рост|");

                                    Razmer = "";
                                }
                                else
                                {
                                    if ((P.sellerText[u + 1].Length >= 4))
                                    {
                                        s = Razmer;
                                        RazmerFind = false;
                                        akamulateRegexLog.Add("# сборка размера закончена - следующая строчка больше 4 символов");

                                        Razmer = "";
                                    }
                                }
                            }
                            else
                            {
                                s = Razmer;
                                RazmerFind = false;
                                akamulateRegexLog.Add("# сборка размера закончена - последняя строчка");

                                Razmer = "";
                            }
                        }



                        if (!RazmerFind)
                        {

                            //В строчке должны быть данные
                            if (s.Length > 1)
                            {
                                // //Добавляем строчку описания в грид
                                // dataGridView3.Rows.Add(s);

                                //Регулярные выражения
                                Regex regex;// = new Regex(@"туп(\w*)", RegexOptions.IgnoreCase);
                                            //Буфферная переменная куда поступают данные после коррекции
                                string resultLine = s;

                                int i = 0;

                                //====================================== Блок замены или стирания ненужной информации ================================================
                                //Производим замену по ключам
                                foreach (ReplaceKeys k in mainForm.Replace_Keys)
                                {
                                    //Если ключ включен, то его исполняем
                                    if (k.RegKey.IsActiv)
                                    {



                                        //Регулярное выражение
                                        regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);

                                        Match M = regex.Match(resultLine);
                                        if (regex.IsMatch(resultLine))
                                        {

                                            //    if (regex.IsMatch(resultLine))
                                            //{
                                            //    P.logRegularExpression.Add(k.RegKey.Value);
                                            //}


                                            //Если режим замены, то заменяем на значение ключа
                                            if (k.Action == 3)
                                            {
                                                //Выполняем замену
                                                resultLine = regex.Replace(resultLine, k.NewValue);

                                                akamulateRegexLog.Add( "# стр." + (u + 1) + " найдено:|" + M.Value + "| Рег.№:" + i + " замена:|" + k.NewValue + "|");
                                            }
                                            //Если режим удаления, то просто вставляем пустое значение
                                            if (k.Action == 4)
                                            {
                                                //Выполняем замену
                                                resultLine = regex.Replace(resultLine, "");

                                                akamulateRegexLog.Add("# стр." + (u + 1) + " найдено:|" + M.Value + "| Рег.№:" + i + " удалена подстрока ");

                                            }

                                            //Если режим удаления, то просто вставляем пустое значение
                                            if (k.Action == 5)
                                            {
                                                akamulateRegexLog.Add("# стр." + (u + 1) + " найдено:|" + M.Value + "| Рег.№:" + i + " строчка заблокирована ");


                                                if (regex.IsMatch(resultLine))
                                                {
                                                    resultLine = "";
                                                }
                                                //Выполняем замену
                                                //resultLine = regex.Replace(resultLine, "");
                                            }
                                        }

                                    }

                                    i++;
                                }


                                //if (resultLine.Length == 0)
                                //{ resultLine = s; }
                                bool reg_line = true;

                                //Проверяем цвет и статус строчки
                                foreach (ColorKeys k in mainForm.Color_Keys)
                                {
                                    //Если ключ активен, то выполняем его
                                    if (k.RegKey.IsActiv)
                                    {
                                        //Регулярное выражение
                                        regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);
                                        //Проверяем вхождение регулярного выражения
                                        bool result = regex.IsMatch(s);//resultLine
                                                                       //Если регулярное выражение сработало
                                        if (result)
                                        {
                                            //Красим строчку в нужный цвет
                                            //dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = k.color;
                                            //Если действие регистрация то добавляем в конструктор поста
                                            if (k.Action == 1)
                                            {
                                                reg_line = false;
                                                //Добавляем к описанию товаров
                                                //stringpost = stringpost + resultLine + "\r\n";
                                            }
                                            break;
                                        }
                                    }
                                }

                                if ((reg_line) && (resultLine.Length > 2))
                                {
                                    P.sellerTextCleen.Add(resultLine);
                                    //Аккамулируем данные для поста
                                    //P.sellerTextCleen.Add(resultLine);

                                    stringpost = stringpost + resultLine + "\r\n";
                                }
                                //Красим добавленную строчку 
                                //dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor;
                            }
                        }

                    }



                    P.logRegularExpression = akamulateRegexLog;


                    string sumPrise = "";
                    string sumSize = "";
                    string sumMaterial = "";

                    string[] RegPrise = new string[5] { "[0-9]{2,7}[ ]?(руб|р|₽)", "ййй", "йййй", "йййй", "йййй" };
                    string[] RegSize = new string[5] { "размер", "ййййй", "йййййй", "ййй", "ййййй" };
                    string[] RegMaterial = new string[5] { "материал", "ййй", "йййй", "ййййй", "йййййй" };

                    int ig = 0;

                    foreach (string Line in P.sellerText)
                    {

                        foreach (string RegEx in RegPrise)
                        {
                            foreach (Match match in Regex.Matches(Line, RegEx, RegexOptions.IgnoreCase))
                            {
                                string price = match.Value;
                                foreach (Match match1 in Regex.Matches(price, "[0-9]{2,7}", RegexOptions.IgnoreCase))
                                {
                                    price = match1.Value;
                                    sumPrise = sumPrise + price + "\r\n";
                                }
                            }
                        }

                        foreach (string RegEx in RegSize)
                        {
                            foreach (Match match in Regex.Matches(Line, RegEx, RegexOptions.IgnoreCase))
                            {
                                string size = Line;// match.Value;
                                sumSize = sumSize + size + "\r\n";
                            }
                        }

                        foreach (string RegEx in RegMaterial)
                        {
                            foreach (Match match in Regex.Matches(Line, RegEx, RegexOptions.IgnoreCase))
                            {
                                string material = Line;//match.Value;
                                sumMaterial = sumMaterial + material + "\r\n";
                            }
                        }
                    }

                    P.Prises = sumPrise;
                    P.Sizes = sumSize;
                    P.Materials = sumMaterial;


                    //if (i == 0)
                    //{ //MessageBox.Show("Внимание такой категории не существует!");
                    //  //break;
                    //}
                    ////============================================================= временное решение ===============================================================================


                    //Пытаемся подобрать категорию по ключам из общей библиотеки категорий
                    autoSelectionCategory(mainForm.mainCategoryList, P, stringpost);

                    //Проверяем была ли на предидущем щаге подобрана категория
                    if (P.CategoryOfProductName == "ВСЕ")
                    {
                        //Подбираем категорию по ключам поставщиков если не подобралась по общей категории
                        autoSelectionCategory(mainForm.providerCategoryList, P, stringpost);
                    }

                    //if (isCat)
                    //{
                    //    isSub = true;
                    //    if (P.SubCategoryOfProductName == "ВСЕ")
                    //    {
                    //        P.SubCategoryOfProductName = "ВСЕ";
                    //    }
                    //}

                    //===============================================================================================================================================================



                    //if (isCat && isSub && isStop)//&& goodsProcessed
                    {

                        mainForm.ProductListForPosting.Add(P);

                        AddToTreeView(mainForm, P, indexTV);

                        RemovedIndexes.Add(index);

                        indexTV++;
                        // ProductListSource.RemoveAt(index);
                        //listBox3.Items.RemoveAt(index);
                    }


                }

                index++;
            }


            return RemovedIndexes;
        }

        /// <summary>
        /// Метод добавления товара в TREEVIEW на форме поста товара
        /// </summary>
        public static void AddToTreeView(MainForm mainForm, Product P, int index)
        {
            bool NodeExistCat = false;
            bool NodeExistSub = false;

            int idNodeExistCat = -1;
            int idNodeExistSub = -1;
            int i = 0;
            int j = 0;

            foreach (TreeNode S in mainForm.treeViewProductForPostBox.Nodes)
            {
                if (S.Text == P.CategoryOfProductName)
                {
                    idNodeExistCat = i;
                    NodeExistCat = true;

                    j = 0;
                    foreach (TreeNode SC in mainForm.treeViewProductForPostBox.Nodes[idNodeExistCat].Nodes)
                    {
                        if (SC.Text == P.SubCategoryOfProductName)
                        {
                            idNodeExistSub = j;
                            NodeExistSub = true;
                            break;
                        }
                        j++;
                    }
                    break;
                }
                i++;
            }


            if (!NodeExistCat)
            {
                idNodeExistCat = mainForm.treeViewProductForPostBox.Nodes.Count;
                idNodeExistSub = 0;

                mainForm.treeViewProductForPostBox.Nodes.Add(P.CategoryOfProductName);

                if (P.SubCategoryOfProductName == "ВСЕ")
                {
                    mainForm.treeViewProductForPostBox.Nodes[mainForm.treeViewProductForPostBox.Nodes.Count - 1].Nodes.Add("ВСЕ");
                }
                else
                {
                    idNodeExistSub++;
                    mainForm.treeViewProductForPostBox.Nodes[mainForm.treeViewProductForPostBox.Nodes.Count - 1].Nodes.Add("ВСЕ");
                    mainForm.treeViewProductForPostBox.Nodes[mainForm.treeViewProductForPostBox.Nodes.Count - 1].Nodes.Add(P.SubCategoryOfProductName);

                }




                NodeExistCat = true;
                NodeExistSub = true;
            }

            if ((NodeExistCat) && (!NodeExistSub))
            {
                idNodeExistSub = mainForm.treeViewProductForPostBox.Nodes[idNodeExistCat].Nodes.Count;

                mainForm.treeViewProductForPostBox.Nodes[idNodeExistCat].Nodes.Add(P.SubCategoryOfProductName);

                NodeExistSub = true;
            }

            if ((NodeExistCat) && (NodeExistSub))
            {
                mainForm.treeViewProductForPostBox.Nodes[idNodeExistCat].Nodes[idNodeExistSub].Nodes.Add(index.ToString());
            }
        }


        /// <summary>
        /// Метод подбора категории товара по ключам из массива категорий
        /// </summary>
        private static void autoSelectionCategory(List<CategoryOfProduct> CategoryList, Product P, string stringpost)
        {
            //========================== Блок с автоподбором категорий ====================================================
            int i = 0;
            int j = 0;
            //Регулярное выражение
            Regex regexCat;
            //Перебираем категории товара
            foreach (CategoryOfProduct c in CategoryList)
            {
                //Флаг обнаружения категории
                bool regincat = false;

                if (P.CategoryOfProductName == c.Name)
                {
                    //listBox1.SelectedIndex = i;
                    regincat = true;
                }



                //Перебираем все ключи привязанные к категории товара
                foreach (Key k in c.Keys)
                {
                    //Если ключ активен, то проверяем его
                    if (k.IsActiv)
                    {
                        //Проверяем регулярное выражение
                        regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);
                        //Аккамулируем результаты поиска 
                        regincat = regexCat.IsMatch(stringpost) || regincat;

                        if (regincat)
                        {
                            Match M = regexCat.Match(stringpost);
                            P.logRegularExpression.Add("# подбор категории:|" + M.Value + "| Рег.:" + k.Value + " категория:|" + c.Name + "|");
                        }
                    }
                }
                //Если категория выбрана, то выделяем ее на форме
                if (regincat)
                {
                    if (P.CategoryOfProductName == "ВСЕ")
                    {
                        P.CategoryOfProductName = c.Name;
                    }
                    //if (!P.HandBlock)
                    //{
                    //    //Выделяем категорию
                    //    listBox1.SelectedIndex = i;
                    //}
                    j = 0;
                    //Перебираем подкатегории выбранной категории
                    foreach (SubCategoryOfProduct s in c.SubCategoty)
                    {
                        //
                        bool reginsubcat = false;

                        //if (P.SubCategoryOfProductName == s.Name)
                        //{
                        //    listBox2.SelectedIndex = j;
                        //}

                        foreach (Key k in s.Keys)
                        {
                            if (k.IsActiv)
                            {
                                regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);

                                reginsubcat = regexCat.IsMatch(stringpost) || reginsubcat;

                                if (reginsubcat)
                                {
                                    Match M = regexCat.Match(stringpost);
                                    P.logRegularExpression.Add("# подбор подкатегории:|" + M.Value + "| Рег.:" + k.Value + " категория:|" + s.Name + "|");
                                }

                            }
                        }

                        if (reginsubcat)
                        {
                            if (P.SubCategoryOfProductName == "ВСЕ")
                            {
                                P.SubCategoryOfProductName = s.Name;
                            }

                            if (!P.HandBlock)
                            {
                                // listBox2.SelectedIndex = j;
                                break;
                            }
                        }
                        j++;
                    }

                    break;
                }
                //Индекс ссылка на категорию товара
                i++;
            }

        }




        #region //Старые блоки с кодом применения регулярных выражений

        ////Проходим по всем строчкам из описания
        //foreach (string s in P.sellerText)//listBox2.SelectedIndex
        //{
        //    //В строчке должны быть данные
        //    if (s.Length > 1)
        //    {
        //        //Регулярные выражения
        //        Regex regex;// = new Regex(@"туп(\w*)", RegexOptions.IgnoreCase);
        //                    //Буфферная переменная куда поступают данные после коррекции
        //        string resultLine = s;



        //        //====================================== Блок замены или стирания ненужной информации ================================================
        //        //Производим замену по ключам
        //        foreach (ReplaceKeys k in Replace_Keys)
        //        {
        //            //Если ключ включен, то его исполняем
        //            if (k.RegKey.IsActiv)
        //            {
        //                //Регулярное выражение
        //                regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);
        //                //Если режим замены, то заменяем на значение ключа
        //                if (k.Action == 3)
        //                {
        //                    //Выполняем замену
        //                    resultLine = regex.Replace(resultLine, k.NewValue);
        //                }
        //                //Если режим удаления, то просто вставляем пустое значение
        //                if (k.Action == 4)
        //                {
        //                    //Выполняем замену
        //                    resultLine = regex.Replace(resultLine, "");
        //                }

        //                if (k.Action == 5)
        //                {
        //                    if (regex.IsMatch(resultLine))
        //                    {
        //                        resultLine = "";
        //                    }
        //                }

        //            }
        //        }






        //        bool reg_line_Green = false;
        //        bool reg_line_Blue = true;

        //        //Проверяем цвет и статус строчки
        //        foreach (ColorKeys k in Color_Keys)
        //        {
        //            //Если ключ активен, то выполняем его
        //            if (k.RegKey.IsActiv)
        //            {
        //                //Регулярное выражение
        //                regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);
        //                //Проверяем вхождение регулярного выражения
        //                bool result = regex.IsMatch(s);//resultLine
        //                                               //Если регулярное выражение сработало
        //                if (result)
        //                {
        //                    //Красим строчку в нужный цвет
        //                    // dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = k.color;
        //                    //Если действие регистрация то добавляем в конструктор поста
        //                    //if (k.Action == 0)
        //                    //{
        //                    //    isStop = false;
        //                    //}

        //                    if ((k.Action == 2))
        //                    {
        //                        reg_line_Green = true;
        //                    }

        //                    if ((k.Action == 1))
        //                    {
        //                        reg_line_Blue = false;
        //                        //Добавляем к описанию товаров
        //                        //stringpost = stringpost + resultLine + "\r\n";
        //                        //P.sellerTextCleen.Add(resultLine);

        //                        ////Добавляем к описанию товаров
        //                        //stringpost = stringpost + resultLine + "\r\n";
        //                    }
        //                    break;
        //                }
        //            }
        //        }




        //        if ((!reg_line_Green && reg_line_Blue) && (resultLine.Length > 2))
        //        {
        //            goodsProcessed = false;
        //        }






        //        if ((true) && (resultLine.Length > 2))//reg_line_Green
        //        {
        //            P.sellerTextCleen.Add(resultLine);
        //            //ProductListSource[listBox3.SelectedIndex].sellerTextCleen.Add(resultLine);
        //            //Аккамулируем данные для поста
        //            stringpost = stringpost + resultLine + "\r\n";
        //        }


        //        ////Аккамулируем данные для поста
        //        //stringpost = stringpost + resultLine + "\r\n";
        //        ////Красим добавленную строчку 
        //        //dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor;
        //    }
        //}

        //foreach (Match match in Regex.Matches(text, @"(\d[.\d]*\d)"))
        //{
        //    string price = match.Value;
        //    int dots = price.Count(i => i.Equals('.'));
        //    if (dots > 1) price = price.Replace(".", "").Insert(price.LastIndexOf(".") - 1, ".");
        //    sum += double.Parse(price, new NumberFormatInfo() { NumberDecimalSeparator = "." });
        //}

















        ////========================== Блок с автоподбором категорий ====================================================
        //int i = 0;
        //int j = 0;
        ////Регулярное выражение
        //Regex regexCat;
        ////Перебираем категории товара
        //foreach (CategoryOfProduct c in mainCategoryList)
        //{
        //    if ((comboBox4.Text == c.Name) || (comboBox4.Text == "ВСЕ"))
        //    {


        //        //Флаг обнаружения категории
        //        bool regincat = false;
        //        //Перебираем все ключи привязанные к категории товара
        //        foreach (Key k in c.Keys)
        //        {
        //            //Если ключ активен, то проверяем его
        //            if (k.IsActiv)
        //            {
        //                //Проверяем регулярное выражение
        //                regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);
        //                //Аккамулируем результаты поиска 
        //                regincat = regexCat.IsMatch(stringpost) || regexCat.IsMatch(P.IDURL) || regincat;
        //            }
        //        }
        //        //Если категория выбрана, то выделяем ее на форме
        //        if (regincat)
        //        {
        //            isCat = true;

        //            if (P.CategoryOfProductName == "ВСЕ")
        //            {
        //                P.CategoryOfProductName = c.Name;
        //            }


        //            j = 0;
        //            //Перебираем подкатегории выбранной категории
        //            foreach (SubCategoryOfProduct s in c.SubCategoty)
        //            {
        //                //
        //                bool reginsubcat = false;
        //                foreach (Key k in s.Keys)
        //                {
        //                    if (k.IsActiv)
        //                    {
        //                        regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);

        //                        reginsubcat = regexCat.IsMatch(stringpost) || regexCat.IsMatch(P.IDURL) || reginsubcat;
        //                    }
        //                }

        //                if (reginsubcat)
        //                {
        //                    isSub = true;

        //                    if (P.SubCategoryOfProductName == "ВСЕ")
        //                    {
        //                        P.SubCategoryOfProductName = s.Name;
        //                    }

        //                    break;
        //                }
        //                j++;
        //            }
        //            break;
        //        }
        //        //Индекс ссылка на категорию товара
        //        i++;

        //    }
        //}


        #endregion 


        public static List<int> ProcessingProducts(MainForm mainForm, int indx)
        {
            int index = 0;
            //int indexTV = 0;
            int indexTV = mainForm.ProductListForPosting.Count;

            List<int> RemovedIndexes = new List<int>();

            Product P = mainForm.productListSource[indx];
            {
                if ((P.IDURL.IndexOf(mainForm.comboBox5.Text) >= 0) || (mainForm.comboBox5.Text == "ВСЕ"))
                {



                    P.sellerTextCleen.Clear();



                    bool isCat = false;
                    bool isSub = false;
                    bool isStop = true;


                    //Буфер с новым описанием
                    string stringpost = "";


                    bool goodsProcessed = true;

                    string Razmer = string.Empty;
                    bool RazmerFind = false;




                    //Проходим по всем строчкам из описания
                    for (int u = 0; u < P.sellerText.Count; u++)//listBox2.SelectedIndex
                    {
                        string s = P.sellerText[u];

                        ////В строчке должны быть данные
                        //if (s.Length > 1)
                        //{
                        //    //Добавляем строчку описания в грид
                        //    dataGridView3.Rows.Add(s);
                        //}

                        if (s.ToLower().IndexOf("размер") >= 0)
                        {
                            RazmerFind = true;
                        }

                        if (RazmerFind)
                        {
                            Razmer = Razmer + " " + s;
                        }

                        if (RazmerFind)
                        {
                            if (u < P.sellerText.Count - 1)
                            {
                                if ((P.sellerText[u + 1].ToLower().IndexOf("рост") >= 0))
                                {
                                    s = Razmer + " " + P.sellerText[u + 1];

                                    u++;

                                    ////В строчке должны быть данные
                                    //if (P.sellerText[u].Length > 1)
                                    //{
                                    //    //Добавляем строчку описания в грид
                                    //    dataGridView3.Rows.Add(P.sellerText[u]);
                                    //}

                                    RazmerFind = false;
                                    Razmer = "";
                                }
                                else
                                {
                                    if ((P.sellerText[u + 1].Length > 4))
                                    {
                                        s = Razmer;
                                        RazmerFind = false;
                                        Razmer = "";
                                    }
                                }
                            }
                            else
                            {
                                s = Razmer;
                                RazmerFind = false;
                                Razmer = "";
                            }
                        }



                        if (!RazmerFind)
                        {

                            //В строчке должны быть данные
                            if (s.Length > 1)
                            {
                                // //Добавляем строчку описания в грид
                                // dataGridView3.Rows.Add(s);

                                //Регулярные выражения
                                Regex regex;// = new Regex(@"туп(\w*)", RegexOptions.IgnoreCase);
                                            //Буфферная переменная куда поступают данные после коррекции
                                string resultLine = s;

                                //====================================== Блок замены или стирания ненужной информации ================================================
                                //Производим замену по ключам
                                foreach (ReplaceKeys k in mainForm.Replace_Keys)
                                {
                                    //Если ключ включен, то его исполняем
                                    if (k.RegKey.IsActiv)
                                    {
                                        //Регулярное выражение
                                        regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);
                                        //Если режим замены, то заменяем на значение ключа
                                        if (k.Action == 3)
                                        {
                                            //Выполняем замену
                                            resultLine = regex.Replace(resultLine, k.NewValue);
                                        }
                                        //Если режим удаления, то просто вставляем пустое значение
                                        if (k.Action == 4)
                                        {
                                            //Выполняем замену
                                            resultLine = regex.Replace(resultLine, "");
                                        }

                                        //Если режим удаления, то просто вставляем пустое значение
                                        if (k.Action == 5)
                                        {
                                            if (regex.IsMatch(resultLine))
                                            {
                                                resultLine = "";
                                            }
                                            //Выполняем замену
                                            //resultLine = regex.Replace(resultLine, "");
                                        }

                                    }
                                }


                                //if (resultLine.Length == 0)
                                //{ resultLine = s; }
                                bool reg_line = true;

                                //Проверяем цвет и статус строчки
                                foreach (ColorKeys k in mainForm.Color_Keys)
                                {
                                    //Если ключ активен, то выполняем его
                                    if (k.RegKey.IsActiv)
                                    {
                                        //Регулярное выражение
                                        regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);
                                        //Проверяем вхождение регулярного выражения
                                        bool result = regex.IsMatch(s);//resultLine
                                                                       //Если регулярное выражение сработало
                                        if (result)
                                        {
                                            //Красим строчку в нужный цвет
                                            //dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = k.color;
                                            //Если действие регистрация то добавляем в конструктор поста
                                            if (k.Action == 1)
                                            {
                                                reg_line = false;
                                                //Добавляем к описанию товаров
                                                //stringpost = stringpost + resultLine + "\r\n";
                                            }
                                            break;
                                        }
                                    }
                                }

                                if ((reg_line) && (resultLine.Length > 2))
                                {
                                    P.sellerTextCleen.Add(resultLine);
                                    //Аккамулируем данные для поста
                                    //P.sellerTextCleen.Add(resultLine);

                                    stringpost = stringpost + resultLine + "\r\n";
                                }
                                //Красим добавленную строчку 
                                //dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor;
                            }
                        }

                    }



                    string sumPrise = "";
                    string sumSize = "";
                    string sumMaterial = "";

                    string[] RegPrise = new string[5] { "[0-9]{2,7}[ ]?(руб|р|₽)", "ййй", "йййй", "йййй", "йййй" };
                    string[] RegSize = new string[5] { "размер", "ййййй", "йййййй", "ййй", "ййййй" };
                    string[] RegMaterial = new string[5] { "материал", "ййй", "йййй", "ййййй", "йййййй" };

                    int ig = 0;

                    foreach (string Line in P.sellerText)
                    {

                        foreach (string RegEx in RegPrise)
                        {
                            foreach (Match match in Regex.Matches(Line, RegEx, RegexOptions.IgnoreCase))
                            {
                                string price = match.Value;
                                foreach (Match match1 in Regex.Matches(price, "[0-9]{2,7}", RegexOptions.IgnoreCase))
                                {
                                    price = match1.Value;
                                    sumPrise = sumPrise + price + "\r\n";
                                }
                            }
                        }

                        foreach (string RegEx in RegSize)
                        {
                            foreach (Match match in Regex.Matches(Line, RegEx, RegexOptions.IgnoreCase))
                            {
                                string size = Line;// match.Value;
                                sumSize = sumSize + size + "\r\n";
                            }
                        }

                        foreach (string RegEx in RegMaterial)
                        {
                            foreach (Match match in Regex.Matches(Line, RegEx, RegexOptions.IgnoreCase))
                            {
                                string material = Line;//match.Value;
                                sumMaterial = sumMaterial + material + "\r\n";
                            }
                        }
                    }

                    P.Prises = sumPrise;
                    P.Sizes = sumSize;
                    P.Materials = sumMaterial;

                    
                    ////============================================================= временное решение ===============================================================================
                    //========================== Блок с автоподбором категорий ====================================================
                    int i = 0;
                    int j = 0;
                    //Регулярное выражение
                    Regex regexCat;
                    //Перебираем категории товара
                    foreach (CategoryOfProduct c in mainForm.mainCategoryList)
                    {
                        //Флаг обнаружения категории
                        bool regincat = false;

                        if (P.CategoryOfProductName == c.Name)
                        {
                            //listBox1.SelectedIndex = i;
                            regincat = true;
                        }



                        //Перебираем все ключи привязанные к категории товара
                        foreach (Key k in c.Keys)
                        {
                            //Если ключ активен, то проверяем его
                            if (k.IsActiv)
                            {
                                //Проверяем регулярное выражение
                                regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);
                                //Аккамулируем результаты поиска 
                                regincat = regexCat.IsMatch(stringpost) || regincat;
                            }
                        }
                        //Если категория выбрана, то выделяем ее на форме
                        if (regincat)
                        {
                            if (P.CategoryOfProductName == "ВСЕ")
                            {
                                P.CategoryOfProductName = c.Name;
                            }
                            //if (!P.HandBlock)
                            //{
                            //    //Выделяем категорию
                            //    listBox1.SelectedIndex = i;
                            //}
                            j = 0;
                            //Перебираем подкатегории выбранной категории
                            foreach (SubCategoryOfProduct s in c.SubCategoty)
                            {
                                //
                                bool reginsubcat = false;

                                //if (P.SubCategoryOfProductName == s.Name)
                                //{
                                //    listBox2.SelectedIndex = j;
                                //}

                                foreach (Key k in s.Keys)
                                {
                                    if (k.IsActiv)
                                    {
                                        regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);

                                        reginsubcat = regexCat.IsMatch(stringpost) || reginsubcat;
                                    }
                                }

                                if (reginsubcat)
                                {
                                    if (P.SubCategoryOfProductName == "ВСЕ")
                                    {
                                        P.SubCategoryOfProductName = s.Name;
                                    }

                                    if (!P.HandBlock)
                                    {
                                        // listBox2.SelectedIndex = j;
                                        break;
                                    }
                                }
                                j++;
                            }

                            break;
                        }
                        //Индекс ссылка на категорию товара
                        i++;
                    }



         

                    //===============================================================================================================================================================



                    // if (isCat && isSub && isStop)//&& goodsProcessed
                    {

                        mainForm.ProductListForPosting.Add(P);

                        AddToTreeView(mainForm, P, indexTV);

                        RemovedIndexes.Add(index);

                        indexTV++;
                        // ProductListSource.RemoveAt(index);
                        //listBox3.Items.RemoveAt(index);
                    }


                }

                index++;
            }


            return RemovedIndexes;
        }


        public static string descriptionProcessing(MainForm mainForm, string s, int u)
        {
            mainForm.logRegexListBox.Items.Add("---------------------------------------------Строчка - " + (u+1) + "----------------------------------------------------");


            //Вспомогательные переменные
            string Razmer = string.Empty;
            bool RazmerFind = false;

            //В строчке должны быть данные
            if (s.Length > 1)
            {
                //Добавляем строчку описания в грид
                mainForm.descriptionSourceDataGridView.Rows.Add(s);
            }

            if (s.ToLower().IndexOf("размер") >= 0)
            {
                RazmerFind = true;
            }

            if (RazmerFind)
            {
                Razmer = Razmer + " " + s;
            }

            if (RazmerFind)
            {
                if (u < mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerText.Count - 1)
                {
                    if ((mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerText[u + 1].ToLower().IndexOf("рост") >= 0))
                    {
                        s = Razmer + " " + mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerText[u + 1];

                        u++;

                        //В строчке должны быть данные
                        if (mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerText[u].Length > 1)
                        {
                            //Добавляем строчку описания в грид
                            mainForm.descriptionSourceDataGridView.Rows.Add(mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerText[u]);
                        }

                        RazmerFind = false;
                        Razmer = "";
                    }
                    else
                    {
                        if ((mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerText[u + 1].Length > 4))
                        {
                            s = Razmer;
                            RazmerFind = false;
                            Razmer = "";
                        }
                    }
                }
                else
                {
                    s = Razmer;
                    RazmerFind = false;
                    Razmer = "";
                }
            }



            if (!RazmerFind)
            {

                //В строчке должны быть данные
                if (s.Length > 1)
                {
                    // //Добавляем строчку описания в грид
                    // dataGridView3.Rows.Add(s);

                    //Регулярные выражения
                    Regex regex;// = new Regex(@"туп(\w*)", RegexOptions.IgnoreCase);
                                //Буфферная переменная куда поступают данные после коррекции
                    string resultLine = s;

                    int i = 0;

                    //====================================== Блок замены или стирания ненужной информации ================================================
                    //Производим замену по ключам
                    foreach (ReplaceKeys k in mainForm.Replace_Keys)
                    {
                        i++;

                        //Если ключ включен, то его исполняем
                        if (k.RegKey.IsActiv)
                        {
                            //Регулярное выражение
                            regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);

                            Match M = regex.Match(resultLine);
                            if (regex.IsMatch(resultLine))
                            {
                                //Если режим замены, то заменяем на значение ключа
                                if (k.Action == 3)
                                {
                                    //Выполняем замену
                                    resultLine = regex.Replace(resultLine, k.NewValue);
                                    mainForm.logRegexListBox.Items.Add("# найдено:|" + M.Value + "| Рег.№:" + i + " замена:|" + k.NewValue+"|");

                                }
                                //Если режим удаления, то просто вставляем пустое значение
                                if (k.Action == 4)
                                {
                                    //Выполняем замену
                                    resultLine = regex.Replace(resultLine, "");
                                    mainForm.logRegexListBox.Items.Add("# найдено:|" + M.Value + "| Рег.№:" + i + " удалена подстрока ");

                                }

                                //Если режим удаления, то просто вставляем пустое значение
                                if (k.Action == 5)
                                {
                                    mainForm.logRegexListBox.Items.Add("# найдено:|" + M.Value + "| Рег.№:" + i + " строчка заблокирована ");


                                    if (regex.IsMatch(resultLine))
                                    {
                                        resultLine = "";
                                    }
                                    //Выполняем замену
                                    //resultLine = regex.Replace(resultLine, "");
                                }
                            }

                        }
                    }


                    //if (resultLine.Length == 0)
                    //{ resultLine = s; }
                    bool reg_line = true;

                    //Проверяем цвет и статус строчки
                    foreach (ColorKeys k in mainForm.Color_Keys)
                    {
                        //Если ключ активен, то выполняем его
                        if (k.RegKey.IsActiv)
                        {
                            //Регулярное выражение
                            regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);
                            //Проверяем вхождение регулярного выражения
                            bool result = regex.IsMatch(s);//resultLine
                                                           //Если регулярное выражение сработало
                            if (result)
                            {
                                //Красим строчку в нужный цвет
                                mainForm.descriptionSourceDataGridView.Rows[mainForm.descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.BackColor = k.color;
                                //Если действие регистрация то добавляем в конструктор поста
                                if (k.Action == 1)
                                {
                                    reg_line = false;
                                    //Добавляем к описанию товаров
                                    //stringpost = stringpost + resultLine + "\r\n";
                                }
                                break;
                            }
                        }
                    }

                    if ((reg_line) && (resultLine.Replace(" ","").Length > 2))
                    {
                        mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].sellerTextCleen.Add(resultLine);

                        //Аккамулируем данные для поста
                        return resultLine + "\r\n";

                        //mainForm.stringpost = mainForm.stringpost + resultLine + "\r\n";
                    }
                    //Красим добавленную строчку 
                    mainForm.descriptionSourceDataGridView.Rows[mainForm.descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = mainForm.descriptionSourceDataGridView.Rows[mainForm.descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.BackColor;
                }
            }

            return "\r\n";

        }

        public static void loadImageInListBox(MainForm mainForm)
        {
            //Чистим картинки в листбоксе
            mainForm.imageListProduct.Images.Clear();
            mainForm.unProcessedProductListView.Items.Clear();

            int ii = 0;
            foreach (string s in mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].FilePath)
            {
                if ((mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].FilePath.Count == 1)
                    && (s == ""))
                {
                    // MessageBox.Show("У данного товара отсутствуют изображения!");
                }
                else
                {
                    try
                    {
                        FileInfo f = new FileInfo(s);
                        if (f.Exists)
                        {

                            mainForm.imageListProduct.Images.Add(new Bitmap(s));
                            mainForm.unProcessedProductListView.Items.Add(new ListViewItem(s, ii));
                            ii++;
                        }
                        else
                        {
                            mainForm.imageNoExist.Add(s);
                        }
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.ToString());
                    }
                }
            }

        }

        public static void categorySelection(MainForm mainForm, string stringpost)
        {
            int i = 0;
            int j = 0;
            //Регулярное выражение
            Regex regexCat;
            //Перебираем категории товара
            foreach (CategoryOfProduct c in mainForm.mainCategoryList)
            {




                //Флаг обнаружения категории
                bool regincat = false;

                if (mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].CategoryOfProductName == c.Name)
                {
                    mainForm.catListBox.SelectedIndex = i;
                    regincat = true;
                }



                //Перебираем все ключи привязанные к категории товара
                foreach (Key k in c.Keys)
                {
                    //Если ключ активен, то проверяем его
                    if (k.IsActiv)
                    {
                        //Проверяем регулярное выражение
                        regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);
                        //Аккамулируем результаты поиска 
                        regincat = regexCat.IsMatch(stringpost) || regincat;
                    }
                }
                //Если категория выбрана, то выделяем ее на форме
                if (regincat)
                {
                    if (mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].CategoryOfProductName == "ВСЕ")
                    {
                        mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].CategoryOfProductName = c.Name;
                    }
                    if (!mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].HandBlock)
                    {
                        //Выделяем категорию
                        mainForm.catListBox.SelectedIndex = i;
                    }
                    j = 0;
                    //Перебираем подкатегории выбранной категории
                    foreach (SubCategoryOfProduct s in c.SubCategoty)
                    {
                        //
                        bool reginsubcat = false;

                        if (mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].SubCategoryOfProductName == s.Name)
                        {
                            mainForm.subCatListBox.SelectedIndex = j;
                        }

                        foreach (Key k in s.Keys)
                        {
                            if (k.IsActiv)
                            {
                                regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);

                                reginsubcat = regexCat.IsMatch(stringpost) || reginsubcat;
                            }
                        }

                        if (reginsubcat)
                        {
                            if (mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].SubCategoryOfProductName == "ВСЕ")
                            {
                                mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].SubCategoryOfProductName = s.Name;
                            }

                            if (!mainForm.productListSource[mainForm.productUnProcessedListBox.SelectedIndex].HandBlock)
                            {
                                mainForm.subCatListBox.SelectedIndex = j;
                                break;
                            }
                        }
                        j++;
                    }

                    break;
                }
                //Индекс ссылка на категорию товара
                i++;
            }


        }


        

        }
    }
