using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Net;
using VKSMM.ModelClasses;//Файл с классами моделей данных
using VKSMM.StuffClasses;//Файл с классами вспомогательных методов
using VKSMM.ThredsCode;//Файл с классами вспомогательных методов

namespace VKSMM
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        //====================================================================================================================================================================
        //                          Блок объявления переменных
        //====================================================================================================================================================================
        /// <summary>
        /// Процесс конвертации XML файлов с товаром во внутренний формат программы
        /// </summary>
        public Thread Thread_Dir_Processing;

        public Thread Thread_Provider_Processing;
        public Thread Thread_Create_XLS_Processing; //Create_XLS

        public Thread t1;
        public Thread t2;
        public Thread t3;
        public Thread t4;

        public string _ProductDBPath = "";
        public string _PhotoPath = "";
        public string _InputPath = "";

        public string _ProviderDir = "";


        public int GUIDReplaceKey = 0;

        public List<CategoryOfProduct> mainCategoryList = new List<CategoryOfProduct>();

        /// <summary>
        /// Список не обработанных товаров. Заполняется при загрузке XLS файлов   
        /// </summary>
        public List<Product> productListSource = new List<Product>();

        public List<Product> ProductListSourceBuffer = new List<Product>();
        

        public List<Product> ProductListForPosting = new List<Product>();

        public List<ReplaceKeys> Replace_Keys = new List<ReplaceKeys>();

        //public List<ReplaceKeys> Addition_Replace_Keys = new List<ReplaceKeys>();

        public List<ColorKeys> Color_Keys = new List<ColorKeys>();
        public List<string> imageNoExist = new List<string>();

        /// <summary>
        /// Путь к месту запуска программы
        /// </summary>
        public string _path = "";





        public int selectedIndexCategory = -1;
        public int selectedIndexSubCategory = -1;
        public int selectedIndexReplace = -1;
        public int selectedIndexColor = -1;

        Regex regex = new Regex(@"туп(\w*)", RegexOptions.IgnoreCase);
        // Regex.IsMatch(email, pattern, RegexOptions.IgnoreCase)

        public List<string> swownProductInListView = new List<string>();

        // Выбрать путь и имя файла в диалоговом окне
        public SaveFileDialog ofd = new SaveFileDialog();


        //====================================================================================================================================================================

        /// <summary>
        /// Действия при загрузке программы
        /// </summary>
        private void MainForm_Load(object sender, EventArgs e)
        {

            //Триггер блокератор блокирующий работу программы после 22 года
            if (DateTime.Now.Year > 2021 && DateTime.Now.Month > 11)
            {
                MessageBox.Show("Ошибка! Обратитесь к разработчику!");
                this.Close();
            }

            //XML документ с настройками программы
            XmlDocument conf_supp = new XmlDocument();
            try
            {
                //Создаем ссылку на файл настроек
                FileInfo F = new FileInfo("ConfigTKSadovod.XML");

                //Копируем старую конфигурацию в директорию копий конфигурацый
                DirectoryInfo d = new DirectoryInfo("CopyOfConfig");
                if (!d.Exists) d.Create();
                F.CopyTo(d.FullName+"\\"+DateTime.Now.Date.ToShortDateString().Replace(".","")+ DateTime.Now.TimeOfDay.ToString().Replace(":","").Replace(".","")+".xml");

                //Сохраняем путь к файлу настроек
                _path = F.FullName;
                //Читаем файл настроек в память
                conf_supp.Load("ConfigTKSadovod.XML");

                //=======================================================================================================
                XmlNode root = conf_supp.DocumentElement;//Получаем доступ к XML файлу настроек
                XmlNodeList nodeList = root.SelectNodes("REPLACE_KEYS"); ;//Вспомогательнаяпеременная

                //Путь к базе с товарами
                _ProductDBPath = F.FullName.Substring(0, F.FullName.LastIndexOf("\\"));

                //Считываем пути к рабочим дирректориям программы
                ConfigReader.readDirFromConfigFile(this, nodeList, root);

                //Считываем ключи замены из конфигурацыы
                ConfigReader.readReplaceKeyFromConfigFile(this, nodeList, root);

                //Считываем цветовые ключи из конфигурацыы
                ConfigReader.readColorKeyFromConfigFile(this, nodeList, root);

                //Считываем категории товаров из конфигурацыонного файла
                ConfigReader.readCategoryProductFromConfigFile(this, nodeList, root);

                //Считываем товары из базы данных обработанных товаров
                ConfigReader.readProductDB(this);

                //Считываем товары из базы данных не обработанных товаров
                ConfigReader.readProductDBUnProcessed(this);

            }
            catch//Если при загрузке конфиг файла случился конфуз, сообщаем пользователю и загружаем по умолчанию
            {
            }

            //Обновляем поставщиков
            Stuff.UpdatePostavshikov(this);
                        
        }

        /// <summary>
        /// Действия при закрытии программы
        /// </summary>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                //Создаем ссылку на старый файл конфигурации
                FileInfo FD = new FileInfo(_path);
                FD.Delete();//Удаляем старый файл конфигурации

                //Объявляем XML настройки
                XmlWriterSettings settings1 = new XmlWriterSettings();
                settings1.Indent = true;
                settings1.NewLineOnAttributes = true;
                settings1.Encoding = Encoding.UTF8;

                //Создаем поток записи в XML файл
                XmlWriter writerConfigXML = XmlWriter.Create(_path, settings1);
                //Записываем заглавие документа
                writerConfigXML.WriteStartDocument();
                //Корневой тег
                writerConfigXML.WriteStartElement("CFG");
                //Сохраняем номер копа
                //writerConfigXML.WriteElementString("AUTHENTIFICATION_STATUS", _AUTHENTIFICATION_STATUS.ToString());

                //Сохраняем пути к рабочим дирректориям программы
                ConfigWriter.writeDirToConfigFile(this, writerConfigXML);

                //Сохраняем ключи замены из конфигурацыы
                ConfigWriter.writeReplaceKeyToConfigFile(this, writerConfigXML);

                //Сохраняем цветовые ключи из конфигурацыы
                ConfigWriter.writeColorKeyToConfigFile(this, writerConfigXML);

                //Сохраняем категории товаров из конфигурацыонного файла
                ConfigWriter.writeCategoryProductToConfigFile(this, writerConfigXML);


                //Закрываем корневой тег
                writerConfigXML.WriteEndElement();
                //Отпускаем поток записи
                writerConfigXML.Close();


                //=======================================================================================================
                //Загружаем товары из базы данных
                //=======================================================================================================
                ConfigWriter.writeProductDB(this);
                //=======================================================================================================


                //=======================================================================================================
                //Загружаем товары из базы данных
                //=======================================================================================================
                ConfigWriter.writeProductDBUnProcessed(this);
                //=======================================================================================================

            }
            catch (Exception e3)
            {
                MessageBox.Show("При сохранение товаров возникла ошибка: " + e3.ToString());
            }

        }

        //====================================================================================================================================================================




        /// <summary>
        /// Действия при выборе не обработанного товара
        /// </summary>
        private void productUnProcessedListBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            //Отчищаем listbox категории и подкатегории
            catListBox.SelectedItems.Clear();
            subCatListBox.SelectedItems.Clear();

            //Если выбран товар то обрабатываем товар
            if (productUnProcessedListBox.SelectedItems.Count>0)
            {
                //========================== Чистим старые данные ================================================================================
                //Стираем данные от постовщика от старого товара
                descriptionSourceDataGridView.Rows.Clear();

                //Отчищаем отредактированное описание выбранного товара. 
                //Оператор мог отредактировать регулярное выражение, по этому
                //описание будет заполнено заново
                productListSource[productUnProcessedListBox.SelectedIndex].sellerTextCleen.Clear();

                //Буфер с новым описанием
                string stringpost = "";

                //Отчищаем лог сработок регулярных выражений
                logRegexListBox.Items.Clear();
                //================================================================================================================================

                //========================== Обновляем данные о товаре ===========================================================================
                //Проходим по всем строчкам из описания
                for (int u = 0;u< productListSource[productUnProcessedListBox.SelectedIndex].sellerText.Count;u++)//listBox2.SelectedIndex
                {
                    //Запускаем процедуру обработки описания товара
                    stringpost += Stuff.descriptionProcessing(this, productListSource[productUnProcessedListBox.SelectedIndex].sellerText[u], u);
                }

                //Добавляем отредактированный пост на форму
                descriptionRegexTextBox.Text = stringpost;

                //Подтягиваем изображения выделенного товара 
                Stuff.loadImageInListBox(this);
                //================================================================================================================================

                //========================== Блок с автоподбором категорий =======================================================================
                Stuff.categorySelection(this, stringpost);

                numericUpDownPrize.Value = productListSource[productUnProcessedListBox.SelectedIndex].prise[0];
                //================================================================================================================================
            }

        }



        private void button9_Click(object sender, EventArgs e)
        {
            string[] Filtr = new string[5];
            Filtr[0] = textBox1.Text.Replace("\r", "").Replace("\n", "");
            Filtr[1] = textBox2.Text.Replace("\r", "").Replace("\n", "");
            Filtr[2] = comboBox3.Text;
            Filtr[3] = GUIDReplaceKey.ToString();
            Filtr[4] = comboBox2.Text;
            GUIDReplaceKey++;

            dataGridView2.Rows.Add(Filtr);


            Key k = new Key();
            k.Value = textBox1.Text.Replace("\r", "").Replace("\n", "");
            k.IsActiv = true;

            ReplaceKeys r = new ReplaceKeys();
            r.Action = Stuff.ActionDecoder(comboBox3.Text);
            r.RegKey = k;
            r.NewValue = textBox2.Text.Replace("\r", "").Replace("\n", "");

            r.GroupValue = comboBox2.Text;

            Replace_Keys.Add(r);
            AddGroup(r);


            //if (comboBox3.Text == "Удалять") dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightCyan;
            //if (comboBox3.Text == "Заменять") dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Green;
            //if (comboBox3.Text == "Дописывать") dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Yellow;
            //if (comboBox3.Text == "Пропускать") dataGridView2.Rows[dataGridView2.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGreen;
        }

    

        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView6.Rows.Clear();

            try
            {
                int u = dataGridView7.SelectedRows[0].Index;
                //int w = dataGridView6.SelectedRows[0].Index;


                if (dataGridView7.SelectedRows[0].Cells[0].Value != null)
                {
                    textBox8.Text = "";
                    dataGridView8.Rows.Clear();

                    int i = 0;

                    foreach (CategoryOfProduct c in mainCategoryList)
                    {
                        dataGridView7.Rows[i].Cells[1].Value = c.SubCategoty.Count;
                        if (c.Name == dataGridView7.SelectedRows[0].Cells[0].Value.ToString())
                        {
                            selectedIndexCategory = i;
                            foreach (SubCategoryOfProduct s in c.SubCategoty)
                            {
                                dataGridView8.Rows.Add(s.Name);
                            }
                            textBox6.Text = c.Name;

                            // break;
                        }
                        i++;
                    }


                    foreach (Key k in mainCategoryList[u].Keys)
                    {
                        dataGridView6.Rows.Add(k.Value);
                    }


                }
            }
            catch { }
        }


        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
           // if (dataGridView5.SelectedRows.Count > 0)
            {
                descriptionSourceDataGridView.Rows.Clear();
                //listBox1.Items.Add(ProductListSource[listBox2.SelectedIndex].sellerText[0]);
                // string[] ssss = new string[ProductListSource[listBox2.SelectedIndex].sellerText.Count]; 
                string stringpost = "";

                int i = 0;
                foreach(Product p in productListSource)
                {
                  //  if(p.IDURL== dataGridView5.SelectedRows[0].Cells[0].Value.ToString())
                    {
                        break;
                    }
                    i++;
                }



                foreach (string s in productListSource[i].sellerText)//listBox2.SelectedIndex
                {
                    descriptionSourceDataGridView.Rows.Add(s);

                    
                    descriptionSourceDataGridView.Rows[descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Red;
                    

                    foreach (DataGridViewRow r in dataGridView2.Rows)
                    {
                        if (s.IndexOf(r.Cells[0].Value.ToString()) >= 0)
                        {
                            descriptionSourceDataGridView.Rows[descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.BackColor = r.DefaultCellStyle.BackColor;
                            break;
                        }
                    }
                    foreach (DataGridViewRow r in dataGridView2.Rows)
                    {
                        if (descriptionSourceDataGridView.Rows[descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.BackColor == Color.LightGreen)
                        {
                            stringpost = stringpost + s + "\r\n";
                            break;
                        }
                    }


                    descriptionSourceDataGridView.Rows[descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = descriptionSourceDataGridView.Rows[descriptionSourceDataGridView.Rows.Count - 1].DefaultCellStyle.BackColor;
                }

                textBox7.Text = stringpost;

                imageListUnProcessedProduct.Images.Clear();
                unProcessedProductListView.Items.Clear();
                //int i = 0;

                //foreach (string s in ProductListSource[listBox2.SelectedIndex].FilePath)
                //{
                //    imageList1.Images.Add(new Bitmap(s));
                //    listView1.Items.Add(new ListViewItem(s, i));
                //    i++;
                //}


            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {

                int u = dataGridView2.SelectedRows[0].Index;
                int i = u;
                //int i = 0;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                dataGridView2.Rows.RemoveAt(i);

                //i = 0;
                ////int j = 0;
                //foreach (ReplaceKeys r in Replace_Keys)
                //{
                //    if (r.RegKey.Value == dataGridView2.SelectedRows[0].Cells[0].Value.ToString())
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;


                Replace_Keys.RemoveAt(i);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string[] Filtr = new string[2];
            Filtr[0] = textBox4.Text;
            Filtr[1] = comboBox1.Text;
            dataGridView4.Rows.Add(Filtr);

            Key k = new Key();
            k.Value = textBox4.Text;
            k.IsActiv = true;

            ColorKeys r = new ColorKeys();
            r.Action = Stuff.ActionDecoder(comboBox1.Text);
            r.RegKey = k;

            if (comboBox1.Text == "Удалять")
                r.color = Color.LightBlue;
            if (comboBox1.Text == "Пропускать")
                r.color = Color.LightGreen;

            Color_Keys.Add(r);


            //if (comboBox1.Text == "Удалять") dataGridView4.Rows[dataGridView4.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightCyan;
            //if (comboBox1.Text == "Заменять") dataGridView4.Rows[dataGridView4.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Green;
            //if (comboBox1.Text == "Дописывать") dataGridView4.Rows[dataGridView4.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Yellow;
            //if (comboBox1.Text == "Пропускать") dataGridView4.Rows[dataGridView4.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGreen;

        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (dataGridView4.SelectedRows.Count > 0)
            {
                int i = 0;
                //int j = 0;
                foreach (DataGridViewRow r in dataGridView4.Rows)
                {
                    if (r.Cells[0].Value == dataGridView4.SelectedRows[0].Cells[0].Value)
                    {
                        break;
                    }
                    i++;
                }
                dataGridView4.Rows.RemoveAt(i);
                Color_Keys.RemoveAt(i);
            }

        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            subCatListBox.Items.Clear();

            if ((productUnProcessedListBox.SelectedIndex >= 0)&&(catListBox.SelectedIndex>=0))
            {
                productListSource[productUnProcessedListBox.SelectedIndex].CategoryOfProductName = catListBox.SelectedItem.ToString();
            }
            
            try
            {
                foreach (SubCategoryOfProduct sub in mainCategoryList[catListBox.SelectedIndex].SubCategoty)
                {
                    subCatListBox.Items.Add(sub.Name);
                }
            }
            catch
            { }
        }

        //Выбор цвета для диференциации
        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView8.SelectedRows[0].Cells[0].Value != null)
                {
                    dataGridView9.Rows.Clear();
                    textBox10.Text = "";
                    

                    int i = 0;

                    foreach (SubCategoryOfProduct c in mainCategoryList[selectedIndexCategory].SubCategoty)
                    {
                        if (c.Name == dataGridView8.SelectedRows[0].Cells[0].Value.ToString())
                        {
                            selectedIndexSubCategory = i;
                            textBox9.Text = c.Name;
                            foreach (Key k in c.Keys)
                            {
                                dataGridView9.Rows.Add(k.Value);
                            }
                            // break;
                        }
                        i++;
                    }

                    //selectedIndexSubCategory

                    //dataGridView6.Rows.Clear();
                    //foreach (Key k in mainCategoryList[selectedIndexCategory].SubCategoty[].Keys)
                    //{
                    //    dataGridView6.Rows.Add(k.Value);
                    //}
                }
            }
            catch { }
        }

        private void dataGridView7_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView9_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ////================================================================================
            //bool GetEmptyString = false;
            //int indexEmptyString = 0;

            //for(int j = 0; j < dataGridView9.Rows.Count-1;j++)
            //{
            //    if(dataGridView9.Rows[j].Cells[0].Value==null)//.ToString().Length<=1)
            //    {
            //        GetEmptyString = true;
            //        indexEmptyString = j;
            //        break;
            //    }
            //    else
            //    {
            //        if (dataGridView9.Rows[j].Cells[0].Value.ToString().Length<=1)
            //        {
            //            GetEmptyString = true;
            //            indexEmptyString = j;
            //            break;
            //        }
            //    }
            //}

            //if(GetEmptyString)
            //{
            //    dataGridView9.Rows.RemoveAt(indexEmptyString);
            //}
            ////================================================================================


            try
            {
                if (dataGridView9.Rows.Count == mainCategoryList[selectedIndexCategory].SubCategoty[selectedIndexSubCategory].Keys.Count + 1)
                {
                    int i = 0;
                    foreach (DataGridViewRow r in dataGridView9.Rows)
                    {
                        if (i == dataGridView9.Rows.Count)
                        {
                        }
                        else 
                        {
                            mainCategoryList[selectedIndexCategory].SubCategoty[selectedIndexSubCategory].Keys[i].Value = r.Cells[0].Value.ToString();
                        }

                        //if(r.Cells[0].Value.ToString()=="Включен")
                        //{
                        //    mainCategoryList[selectedIndexCategory].Keys[i].IsActiv = true;
                        //}   
                        //else
                        //{
                        //    mainCategoryList[selectedIndexCategory].Keys[i].IsActiv = false;
                        //}

                        i++;
                    }
                }
                else
                {
                    if (dataGridView9.Rows[dataGridView9.Rows.Count - 2].Cells[0].Value != null)
                    {
                        Key newKey = new Key();
                        newKey.IsActiv = true;
                        newKey.Value = dataGridView9.Rows[dataGridView9.Rows.Count - 2].Cells[0].Value.ToString();
                        //newSubCP.Name = dataGridView8.Rows[dataGridView8.Rows.Count - 2].Cells[0].Value.ToString();
                        mainCategoryList[selectedIndexCategory].SubCategoty[selectedIndexSubCategory].Keys.Add(newKey);
                    }
                }
            }
            catch { }




        }


        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            //    //if (dataGridView4.SelectedRows.Count > 0)
            //    //{
            //    //    int i = 0;
            //    //    foreach (DataGridViewRow r in dataGridView4.Rows)
            //    //    {
            //    //        if (r.Cells[0].Value == dataGridView4.SelectedRows[0].Cells[0].Value)
            //    //        {
            //    //            break;
            //    //        }
            //    //        i++;
            //    //    }

            //        Color_Keys[selectedIndexColor].RegKey.Value = dataGridView4.SelectedRows[0].Cells[0].Value.ToString();
            //        //Color_Keys[i].NewValue = dataGridView4.SelectedRows[0].Cells[1].Value.ToString();
            //        Color_Keys[selectedIndexColor].Action = ActionDecoder(dataGridView4.SelectedRows[0].Cells[1].Value.ToString());
            //   // }
            //}
            //catch { }

        }



        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    int i = 0;
                    //foreach (DataGridViewRow r in dataGridView2.Rows)
                    //{
                    //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                    //    {
                    //        selectedIndexReplace = i;

                    //        textBox1.Text = Replace_Keys[selectedIndexReplace].RegKey.Value;
                    //        textBox2.Text = Replace_Keys[selectedIndexReplace].NewValue;
                    //        comboBox3.Text = ActionCoder(Replace_Keys[selectedIndexReplace].Action);
                    //        comboBox2.Text = Replace_Keys[selectedIndexReplace].GroupValue;

                    //        break;
                    //    }
                    //    i++;
                    //}

                    i = 0;
                    //int j = 0;
                    foreach (ReplaceKeys r in Replace_Keys)
                    {
                        if (r.RegKey.Value == dataGridView2.SelectedRows[0].Cells[0].Value.ToString())
                        {
                            break;
                        }
                        i++;
                    }
                    selectedIndexReplace = i;


                    //selectedIndexReplace = Convert.ToInt32(dataGridView2.SelectedRows[0].Cells[3].Value);

                    textBox1.Text = Replace_Keys[selectedIndexReplace].RegKey.Value;
                    textBox2.Text = Replace_Keys[selectedIndexReplace].NewValue;
                    comboBox3.Text = Stuff.ActionCoder(Replace_Keys[selectedIndexReplace].Action);
                    comboBox2.Text = Replace_Keys[selectedIndexReplace].GroupValue;
                }
            }
            catch
            { }

        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.SelectedRows.Count > 0)
            {
                int i = 0;
                foreach (DataGridViewRow r in dataGridView4.Rows)
                {
                    if (r.Cells[0].Value == dataGridView4.SelectedRows[0].Cells[0].Value)
                    {
                        selectedIndexColor = i;

                        textBox4.Text = Color_Keys[selectedIndexColor].RegKey.Value;
                        comboBox1.Text = Stuff.ActionCoder(Color_Keys[selectedIndexColor].Action);


                        break;
                    }
                    i++;
                }

            }

        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            if ((dataGridView2.SelectedRows.Count > 0) )
            {
                int u = dataGridView2.SelectedRows[0].Index;
                int i = u;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //dataGridView2.Rows[selectedIndexReplace].GroupValue = comboBox2.Text;
                dataGridView2.Rows[i].Cells[0].Value = textBox1.Text.Replace("\r", "").Replace("\n", "");
                dataGridView2.Rows[i].Cells[1].Value = textBox2.Text.Replace("\r", "").Replace("\n", "");
                dataGridView2.Rows[i].Cells[2].Value = comboBox3.Text;
                dataGridView2.Rows[i].Cells[4].Value = comboBox2.Text;
                //i = 0;
                ////int j = 0;
                //foreach (ReplaceKeys r in Replace_Keys)
                //{
                //    if (r.RegKey.Value == dataGridView2.SelectedRows[0].Cells[0].Value.ToString())
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //dataGridView5.Rows[i].Cells[0].Value = textBox1.Text;
                //dataGridView5.Rows[i].Cells[1].Value = textBox2.Text;
                //dataGridView5.Rows[i].Cells[2].Value = comboBox3.Text;



                Replace_Keys[i].GroupValue = comboBox2.Text;
                Replace_Keys[i].RegKey.Value = textBox1.Text.Replace("\r", "").Replace("\n", "");
                Replace_Keys[i].NewValue = textBox2.Text.Replace("\r", "").Replace("\n", "");
                Replace_Keys[i].Action = Stuff.ActionDecoder(comboBox3.Text);

                AddGroup(Replace_Keys[i]);
            }

        }

        private void button14_Click(object sender, EventArgs e)
        {
            Color_Keys[selectedIndexColor].RegKey.Value = textBox4.Text;
            Color_Keys[selectedIndexColor].Action = Stuff.ActionDecoder(comboBox1.Text);

            dataGridView4.Rows[selectedIndexColor].Cells[0].Value = textBox4.Text;
            dataGridView4.Rows[selectedIndexColor].Cells[1].Value = comboBox1.Text;

        }

        private void PublicationButton_Click(object sender, EventArgs e)
        {

            try
            {
                treeView1.SelectedNode = treeView1.Nodes[0];
            }
            catch
            { }






            List<int> RemovedIndexes = new List<int>();


            RemovedIndexes = Stuff.ProcessingProducts(this, productUnProcessedListBox.SelectedIndex);


            for (int iq = RemovedIndexes.Count - 1; iq >= 0; iq--)
            {
                productUnProcessedListBox.Items.RemoveAt(RemovedIndexes[iq]);
                productListSource.RemoveAt(RemovedIndexes[iq]);
            }

            Stuff.UpdatePostavshikov(this);

            productUnProcessedListBox.SelectedItems.Clear();

            descriptionRegexTextBox.Text = "";
            unProcessedProductListView.Items.Clear();
            numericUpDownPrize.Value = 0;
            descriptionSourceDataGridView.Rows.Clear();

           
        }

        public void AddToTreeView(Product P, int index)
        {
            bool NodeExistCat = false;
            bool NodeExistSub = false;

            int idNodeExistCat = -1;
            int idNodeExistSub = -1;
            int i = 0;
            int j = 0;

            foreach (TreeNode S in treeView1.Nodes)
            {
                if (S.Text == P.CategoryOfProductName)
                {
                    idNodeExistCat = i;
                    NodeExistCat = true;

                    j = 0;
                    foreach (TreeNode SC in treeView1.Nodes[idNodeExistCat].Nodes)
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
                idNodeExistCat = treeView1.Nodes.Count;
                idNodeExistSub = 0;

                treeView1.Nodes.Add(P.CategoryOfProductName);

                if (P.SubCategoryOfProductName == "ВСЕ")
                {
                    treeView1.Nodes[treeView1.Nodes.Count - 1].Nodes.Add("ВСЕ");
                }
                else
                {
                    idNodeExistSub++;
                    treeView1.Nodes[treeView1.Nodes.Count - 1].Nodes.Add("ВСЕ");
                    treeView1.Nodes[treeView1.Nodes.Count - 1].Nodes.Add(P.SubCategoryOfProductName);

                }




                NodeExistCat = true;
                NodeExistSub = true;
            }

            if ((NodeExistCat) && (!NodeExistSub))
            {
                idNodeExistSub = treeView1.Nodes[idNodeExistCat].Nodes.Count;

                treeView1.Nodes[idNodeExistCat].Nodes.Add(P.SubCategoryOfProductName);

                NodeExistSub = true;
            }

            if ((NodeExistCat) && (NodeExistSub))
            {
                treeView1.Nodes[idNodeExistCat].Nodes[idNodeExistSub].Nodes.Add(index.ToString());
            }
        }

        public void AddGroup(ReplaceKeys R)
        {
            bool get = true;
            foreach(string s in comboBox2.Items)
            {
                if(s==R.GroupValue)
                {
                    get = false;
                }
            }

            if(get)
            {
                comboBox2.Items.Add(R.GroupValue);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                treeView1.SelectedNode = treeView1.Nodes[0];
            }
            catch
            { }




            productUnProcessedListBox.SelectedItems.Clear();

            descriptionRegexTextBox.Text = "";
            unProcessedProductListView.Items.Clear();
            numericUpDownPrize.Value = 0;
            descriptionSourceDataGridView.Rows.Clear();


            List<int> RemovedIndexes = new List<int>();


            RemovedIndexes = Stuff.ProcessingProductsAll(this);


            for (int iq = RemovedIndexes.Count-1; iq>=0;iq-- )
            {
                productUnProcessedListBox.Items.RemoveAt(RemovedIndexes[iq]);
                productListSource.RemoveAt(RemovedIndexes[iq]);
            }

            Stuff.UpdatePostavshikov(this);
        }




        //public class Solution
        //{
        //    public int[] TwoSum(int[] nums, int target)
        //    {

        //        for (int i = 0; i < nums.Length; i++)
        //        {
        //            if (nums[i] < target)
        //            {
        //                for (int j = 0; j < nums.Length; j++)
        //                {
        //                    if (nums[j] != nums[i] && nums[j] < target)
        //                    {
        //                        if (nums[i] + nums[j] == target)
        //                        {
        //                            return new int[] { i, j };
        //                        }
        //                    }

        //                }
        //            }
        //        }
        //    }
        //}




        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if(treeView1.SelectedNode.Level==2)
            {

                button29.Enabled = false;
                button30.Enabled = false;


                textBox7.Text = "";
                textBox5.Text = "";
                listBox5.Items.Clear();
                listBox4.Items.Clear();
                int ic = 0;

                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);
                int indexcat = -1;
                int indexSUB = -1;

                try
                {
                    foreach (CategoryOfProduct cat in mainCategoryList)
                    {
                        listBox5.Items.Add(cat.Name);

                        if(cat.Name==ProductListForPosting[indexOfProduct].CategoryOfProductName)
                        {
                            indexcat = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }
                ic = 0;
                try
                {
                    foreach (SubCategoryOfProduct SUB in mainCategoryList[indexcat].SubCategoty)
                    {
                        listBox4.Items.Add(SUB.Name);

                        if (SUB.Name == ProductListForPosting[indexOfProduct].SubCategoryOfProductName)
                        {
                            indexSUB = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }


                listBox5.SelectedIndex = indexcat;
                listBox4.SelectedIndex = indexSUB;



                //Чистим картинки в листбоксе
                imageListUnProcessedProduct.Images.Clear();
                listView2.Items.Clear();

                int ii = 0;
                foreach (string s in ProductListForPosting[indexOfProduct].FilePath)
                {
                    try
                    {
                        if (s.Length > 3)
                        {
                            imageListUnProcessedProduct.Images.Add(new Bitmap(s));
                            listView2.Items.Add(new ListViewItem(s, ii));
                            ii++;
                        }
                    }
                    catch { }
                }

                string sout = "";
                foreach (string sss in ProductListForPosting[indexOfProduct].sellerTextCleen)
                {
                    sout = sout + sss + "\r\n";
                }
                textBox7.Text = sout;
                sout = "";
                foreach (string sss in ProductListForPosting[indexOfProduct].sellerText)
                {
                    sout = sout + sss + "\r\n";
                }
                textBox11.Text = sout;


                textBox5.Text = ProductListForPosting[indexOfProduct].Prises;
            }


            if (treeView1.SelectedNode.Level == 1)
            {

                button29.Enabled = false;
                button30.Enabled = true;

                swownProductInListView = new List<string>();



                textBox7.Text = "";
                textBox5.Text = "";
                listBox5.Items.Clear();
                listBox4.Items.Clear();
                int ic = 0;

                label16.Text = treeView1.SelectedNode.Parent.Text +"/"+ treeView1.SelectedNode.Text;


                int indexcat = -1;
                int indexSUB = -1;

                try
                {
                    foreach (CategoryOfProduct cat in mainCategoryList)
                    {
                        listBox5.Items.Add(cat.Name);

                        if (cat.Name == treeView1.SelectedNode.Parent.Text)
                        {
                            indexcat = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }
                ic = 0;
                try
                {
                    foreach (SubCategoryOfProduct SUB in mainCategoryList[indexcat].SubCategoty)
                    {
                        listBox4.Items.Add(SUB.Name);

                        if (SUB.Name == treeView1.SelectedNode.Text)
                        {
                            indexSUB = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }


                listBox5.SelectedIndex = indexcat;
                listBox4.SelectedIndex = indexSUB;



                //Чистим картинки в листбоксе
                imageListUnProcessedProduct.Images.Clear();
                listView2.Items.Clear();

                int ii = 0;
                foreach (TreeNode tns in treeView1.SelectedNode.Nodes)
                {
                    int indexOfProduct = Convert.ToInt32(tns.Text);
                   
                   // foreach (string s in ProductListForPosting[indexOfProduct].FilePath)
                    {
                        try
                        {
                            string s = ProductListForPosting[indexOfProduct].FilePath[0];
                            if (s.Length > 3)
                            {


                                swownProductInListView.Add(s);

                                if (ii < 100)
                                {
                                    imageListUnProcessedProduct.Images.Add(new Bitmap(s));
                                    listView2.Items.Add(new ListViewItem(s, ii));
                                }

                                ii++;

                                //imageList1.Images.Add(new Bitmap(s));
                                //listView2.Items.Add(new ListViewItem(s, ii));
                                //ii++;
                            }
                        }
                        catch
                        { }
                    }
                }


                int countShowProduct = (int)(swownProductInListView.Count / 100) + 1;

                if (countShowProduct == 1)
                {
                    button30.Enabled = false;
                }

                button29.Text = "<- 1";
                button30.Text = countShowProduct + " ->";

                //string sout = "";
                //foreach (string sss in ProductListForPosting[indexOfProduct].sellerTextCleen)
                //{
                //    sout = sout + sss + "\r\n";
                //}
                //textBox7.Text = sout;


                // textBox5.Text = ProductListForPosting[indexOfProduct].Prises;
            }


            if (treeView1.SelectedNode.Level == 0)
            {

                button29.Enabled = false;
                button30.Enabled = true;

                swownProductInListView = new List<string>();



                textBox7.Text = "";
                textBox5.Text = "";
                listBox5.Items.Clear();
                listBox4.Items.Clear();

                label16.Text = treeView1.SelectedNode.Text;

                int ic = 0;

               // int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);
                int indexcat = -1;
                int indexSUB = -1;

                try
                {
                    foreach (CategoryOfProduct cat in mainCategoryList)
                    {
                        listBox5.Items.Add(cat.Name);

                        if (cat.Name == treeView1.SelectedNode.Text)
                        {
                            indexcat = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }
                //ic = 0;
                //try
                //{
                //    foreach (SubCategoryOfProduct SUB in mainCategoryList[indexcat].SubCategoty)
                //    {
                //        listBox4.Items.Add(SUB.Name);

                //        if (SUB.Name == ProductListForPosting[indexOfProduct].SubCategoryOfProductName)
                //        {
                //            indexSUB = ic;
                //        }
                //        ic++;
                //    }
                //}
                //catch
                //{ }


                listBox5.SelectedIndex = indexcat;
                //listBox4.SelectedIndex = indexSUB;



                //Чистим картинки в листбоксе
                imageListUnProcessedProduct.Images.Clear();
                listView2.Items.Clear();
                int ii = 0;

                foreach (TreeNode cns in treeView1.SelectedNode.Nodes)
                {

                    foreach (TreeNode tns in cns.Nodes)
                    {
                        try
                        {
                            int indexOfProduct = Convert.ToInt32(tns.Text);

                            if (ProductListForPosting[indexOfProduct].FilePath[0].Length > 3)
                            {
                                swownProductInListView.Add(ProductListForPosting[indexOfProduct].FilePath[0]);

                                if (ii < 100)
                                {
                                    imageListUnProcessedProduct.Images.Add(new Bitmap(ProductListForPosting[indexOfProduct].FilePath[0]));
                                    listView2.Items.Add(new ListViewItem(ProductListForPosting[indexOfProduct].FilePath[0], ii));
                                }
                                
                                ii++;
                            }
                        }
                        catch
                        { }
                    }
                }


                int countShowProduct = (int)(swownProductInListView.Count/100)+1;

                if(countShowProduct == 1)
                {
                    button30.Enabled = false;
                }

                button29.Text = "<- 1";
                button30.Text = countShowProduct + " ->";
                //string sout = "";
                //foreach (string sss in ProductListForPosting[indexOfProduct].sellerTextCleen)
                //{
                //    sout = sout + sss + "\r\n";
                //}
                //textBox7.Text = sout;


                //textBox5.Text = ProductListForPosting[indexOfProduct].Prises;
            }

        }


        private void listBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox4.Items.Clear();

            //if ((listBox3.SelectedIndex >= 0) && (listBox1.SelectedIndex >= 0))
            //{
            //    ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = listBox1.SelectedItem.ToString();
            //}

            try
            {
                foreach (SubCategoryOfProduct sub in mainCategoryList[listBox5.SelectedIndex].SubCategoty)
                {
                    listBox4.Items.Add(sub.Name);
                }
            }
            catch
            { }



            //listBox4.Items.Clear();

            //if ((listBox3.SelectedIndex >= 0) && (listBox1.SelectedIndex >= 0))
            //{
            //    ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = listBox1.SelectedItem.ToString();
            //}

            //try
            //{
            //    foreach (SubCategoryOfProduct sub in mainCategoryList[listBox1.SelectedIndex].SubCategoty)
            //    {
            //        listBox2.Items.Add(sub.Name);
            //    }
            //}
            //catch
            //{ }






            //if (treeView1.SelectedNode.Level == 2)
            //{

            //    int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);


            //    ProductListForPosting[indexOfProduct].CategoryOfProductName = listBox5.SelectedItem.ToString();
            //}

        }


        private void button5_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Level == 2)
            {

                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);



                Product double_produkt = new Product();


                double_produkt.CategoryOfProductName = ProductListForPosting[indexOfProduct].CategoryOfProductName;
                double_produkt.datePost = ProductListForPosting[indexOfProduct].datePost;

                foreach (string sss in ProductListForPosting[indexOfProduct].FilePath)
                {
                    double_produkt.FilePath.Add(sss);
                }

                double_produkt.IDURL = ProductListForPosting[indexOfProduct].IDURL;

                for (int i = 0; i < ProductListForPosting[indexOfProduct].prise.Length; i++)
                {
                    double_produkt.prise[i] = ProductListForPosting[indexOfProduct].prise[i];
                }

                foreach (string sss in ProductListForPosting[indexOfProduct].sellerText)
                {
                    double_produkt.sellerText.Add(sss);
                }

                foreach (string sss in ProductListForPosting[indexOfProduct].sellerTextCleen)
                {
                    double_produkt.sellerTextCleen.Add(sss);
                }

                double_produkt.SubCategoryOfProductName = ProductListForPosting[indexOfProduct].SubCategoryOfProductName;

                foreach (string sss in ProductListForPosting[indexOfProduct].URLPhoto)
                {
                    double_produkt.URLPhoto.Add(sss);
                }


                treeView1.SelectedNode.Parent.Nodes.Add(ProductListForPosting.Count.ToString());
                ProductListForPosting.Add(double_produkt);

            }

        }



        private void button6_Click(object sender, EventArgs e)
        {
            // Задаем расширение имени файла по умолчанию (открывается папка с программой)
            ofd.DefaultExt = "*.xls;*.xlsx";
            // Задаем строку фильтра имен файлов, которая определяет варианты
            ofd.Filter = "файл Excel (*.xlsx)|*.xlsx";
            // Задаем заголовок диалогового окна
            ofd.Title = "Сохранение файла для поста";
            if ((ofd.ShowDialog() == DialogResult.OK)) // если файл БД не выбран -> Выход
            {

                Thread_Create_XLS_Processing = new Thread(Thread_Create_XLS_Processing_Code);
                Thread_Create_XLS_Processing.Start();
            }
        }


        void Thread_Create_XLS_Processing_Code()
        { 
            Stuff.CreateExcel(this, ofd);
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (treeView1.SelectedNode.Level == 2)
            //{

            //    int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);

            //    try
            //    {
            //        ProductListForPosting[indexOfProduct].SubCategoryOfProductName = listBox4.SelectedItem.ToString();
            //    }
            //    catch
            //    {

            //    }

            //    treeView1.Nodes.Remove(treeView1.SelectedNode);

            //    AddToTreeView(ProductListForPosting[indexOfProduct], indexOfProduct);
            //}
        }

        private void textBox7_KeyUp(object sender, KeyEventArgs e)
        {
            if (treeView1.SelectedNode.Level == 2)
            {

                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);

                ProductListForPosting[indexOfProduct].sellerTextCleen.Clear();
                foreach (string s in textBox7.Text.Split('\n'))
                {
                    ProductListForPosting[indexOfProduct].sellerTextCleen.Add(s);
                }


                // numericUpDown1.Value = ProductListForPosting[indexOfProduct].prise[0];
            }

        }

        private void numericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
            if (treeView1.SelectedNode.Level == 2)
            {

                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);


                ProductListForPosting[indexOfProduct].Prises = textBox5.Text;
            }

        }

        private void listView2_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode==Keys.Delete)
            {
                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);

                ProductListForPosting[indexOfProduct].URLPhoto.RemoveAt(listView2.SelectedIndices[0]);
                ProductListForPosting[indexOfProduct].FilePath.RemoveAt(listView2.SelectedIndices[0]);


                listView2.Items.Remove(listView2.SelectedItems[0]);
            }
        }

        private void numericUpDown1_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Level == 2)
            {

                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);


                ProductListForPosting[indexOfProduct].Prises = textBox5.Text;
            }

        }

        private void listBox4_Click(object sender, EventArgs e)
        {
            if (((treeView1.SelectedNode.Level == 0)||(treeView1.SelectedNode.Level == 1)) && (listView2.SelectedItems.Count > 0))//
            {

                string Line1 = listBox5.SelectedItem.ToString();
                string Line2 = listBox4.SelectedItem.ToString();
                int iii = listView2.SelectedItems.Count - 1;

                for(int ii= iii; ii>=0 ;ii--)
                //while(listView2.SelectedItems.Count>0)
                {



                    int indexOfProduct = -1;
                    int i = 0;
                    foreach (Product P in ProductListForPosting)
                    {
                        if (listView2.SelectedItems[ii].Text == P.FilePath[0])
                        {
                            indexOfProduct = i;
                            break;
                        }
                        i++;
                    }



                    try
                    {
                        ProductListForPosting[indexOfProduct].CategoryOfProductName = Line1;
                        ProductListForPosting[indexOfProduct].SubCategoryOfProductName = Line2;
                        listView2.Items.RemoveAt(listView2.SelectedItems[ii].Index);
                    }
                    catch
                    {

                    }
                    bool gt = false;

                    foreach (TreeNode t in treeView1.Nodes)
                    {
                        foreach (TreeNode tt in t.Nodes)
                        {
                            int ht = 0;
                            foreach (TreeNode ttt in tt.Nodes)
                            {
                                if (ttt.Text == indexOfProduct.ToString())
                                {
                                    gt = true;
                                    break;
                                }
                                ht++;
                            }

                            if (gt)
                            {
                                tt.Nodes.RemoveAt(ht);
                                break;
                            }
                        }
                        if (gt)
                        {
                            break;
                        }
                    }

                    //treeView1.Nodes.Remove(treeView1.SelectedNode);

                    AddToTreeView(ProductListForPosting[indexOfProduct], indexOfProduct);
                }
            }
            else
            {



                if (treeView1.SelectedNode.Level == 2)
                {

                    int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);

                    try
                    {
                        ProductListForPosting[indexOfProduct].CategoryOfProductName = listBox5.SelectedItem.ToString();
                        ProductListForPosting[indexOfProduct].SubCategoryOfProductName = listBox4.SelectedItem.ToString();
                    }
                    catch
                    {

                    }

                    treeView1.Nodes.Remove(treeView1.SelectedNode);

                    AddToTreeView(ProductListForPosting[indexOfProduct], indexOfProduct);
                }
            }

        }

        private void textBox5_KeyUp(object sender, KeyEventArgs e)
        {
            if (treeView1.SelectedNode.Level == 2)
            {

                int indexOfProduct = Convert.ToInt32(treeView1.SelectedNode.Text);


                ProductListForPosting[indexOfProduct].Prises = textBox5.Text;
            }

        }

        private void listBox2_MouseUp(object sender, MouseEventArgs e)
        {
            if ((productUnProcessedListBox.SelectedIndex >= 0) && (subCatListBox.SelectedIndex >= 0))
            {
                // ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = listBox1.SelectedItem.ToString();

                productListSource[productUnProcessedListBox.SelectedIndex].SubCategoryOfProductName = subCatListBox.SelectedItem.ToString();

                productListSource[productUnProcessedListBox.SelectedIndex].HandBlock = true;
            }

        }


        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
          

        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();

            if ("ВСЕ" == comboBox2.Text)
            {
                foreach (ReplaceKeys RK in Replace_Keys)
                {
                    {
                        //=========== Конструктор для добавления в датагрид =============
                        //Строка для добавления на грид
                        string[] Filtr = new string[4];
                        //Ключ замены регулярное выражение
                        Filtr[0] = RK.RegKey.Value;
                        //Filtr[0] = PL.ChildNodes[0].InnerText;
                        //Значение замены
                        Filtr[1] = RK.NewValue;
                        //Действие
                        Filtr[2] = Stuff.ActionCoder(RK.Action);

                        Filtr[3] = GUIDReplaceKey.ToString();
                        GUIDReplaceKey++;

                        //Добавляем правило в ГРИД
                        dataGridView2.Rows.Add(Filtr);
                        //================================================================
                    }

                }

            }
            else
            {
                List<ReplaceKeys> M = new List<ReplaceKeys>();

                foreach (ReplaceKeys RK in Replace_Keys)
                {
                    if (RK.GroupValue == comboBox2.Text)
                    {
                        M.Add(RK);
                    }

                }

                foreach(ReplaceKeys RK in M)
                {
                    //=========== Конструктор для добавления в датагрид =============
                    //Строка для добавления на грид
                    string[] Filtr = new string[4];
                    //Ключ замены регулярное выражение
                    Filtr[0] = RK.RegKey.Value;
                    //Filtr[0] = PL.ChildNodes[0].InnerText;
                    //Значение замены
                    Filtr[1] = RK.NewValue;
                    //Действие
                    Filtr[2] = Stuff.ActionCoder(RK.Action);

                    Filtr[3] = GUIDReplaceKey.ToString();
                    GUIDReplaceKey++;

                    //Добавляем правило в ГРИД
                    dataGridView2.Rows.Add(Filtr);
                    //================================================================

                }
            }

        }


        private void button4_Click(object sender, EventArgs e)
        {
            if ((dataGridView2.SelectedRows.Count > 0)&& (dataGridView2.Rows.Count == Replace_Keys.Count))
            {
                int u = dataGridView2.SelectedRows[0].Index;
                int i = u;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                if (i > 0)
                {

                    string[] s = new string[5];
                    s[0] = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    s[1] = dataGridView2.Rows[i].Cells[1].Value.ToString();
                    s[2] = dataGridView2.Rows[i].Cells[2].Value.ToString();
                    s[3] = dataGridView2.Rows[i].Cells[3].Value.ToString();
                    s[4] = dataGridView2.Rows[i].Cells[4].Value.ToString();

                    dataGridView2.Rows.Insert(i - 1, s);
                    dataGridView2.Rows.RemoveAt(i + 1);
                    // dataGridView2.ClearSelection();

                    //i = 0;
                    ////int j = 0;
                    //foreach (ReplaceKeys r in Replace_Keys)
                    //{
                    //    if (r.RegKey.Value == dataGridView2.Rows[u].Cells[0].Value.ToString())
                    //    {
                    //        break;
                    //    }
                    //    i++;
                    //}
                    //int i = selectedIndexReplace;

                    Replace_Keys.Insert(i - 1, Replace_Keys[i]);
                    Replace_Keys.RemoveAt(i + 1);

                    dataGridView2.ClearSelection();

                    if (u - 1 >= 0)
                    {
                        dataGridView2.Rows[u - 1].Selected = true;
                    }
                    else
                    {

                        dataGridView2.Rows[0].Selected = true;

                    }
                }
            }


            listBox6.Items.Clear();
            foreach (ReplaceKeys r in Replace_Keys)
            {
                listBox6.Items.Add(r.RegKey.Value);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if ((dataGridView2.SelectedRows.Count > 0)&&(dataGridView2.Rows.Count == Replace_Keys.Count))
            {
                int u = dataGridView2.SelectedRows[0].Index;
                int i = u;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                if (i < dataGridView2.Rows.Count-1)
                {

                    string[] s = new string[5];
                    s[0] = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    s[1] = dataGridView2.Rows[i].Cells[1].Value.ToString();
                    s[2] = dataGridView2.Rows[i].Cells[2].Value.ToString();
                    s[3] = dataGridView2.Rows[i].Cells[3].Value.ToString();
                    s[4] = dataGridView2.Rows[i].Cells[4].Value.ToString();

                    dataGridView2.Rows.Insert(i + 2, s);
                    dataGridView2.Rows.RemoveAt(i);
                    // dataGridView2.ClearSelection();

                    //i = 0;
                    ////int j = 0;
                    //foreach (ReplaceKeys r in Replace_Keys)
                    //{
                    //    if (r.RegKey.Value == dataGridView2.Rows[u].Cells[0].Value.ToString())
                    //    {
                    //        break;
                    //    }
                    //    i++;
                    //}
                    //int i = selectedIndexReplace;

                    Replace_Keys.Insert(i + 2, Replace_Keys[i ]);
                    Replace_Keys.RemoveAt(i  );


                    dataGridView2.ClearSelection();

                    if (u + 1 < dataGridView2.Rows.Count)
                    {
                        dataGridView2.Rows[u + 1].Selected = true;
                    }
                    else
                    {
                        dataGridView2.Rows[dataGridView2.Rows.Count-1].Selected = true;
                    }
                }
            }
            listBox6.Items.Clear();
            foreach (ReplaceKeys r in Replace_Keys)
            {
                listBox6.Items.Add(r.RegKey.Value);
            }
        }


        /// <summary>
        /// Действия при нажатии кнопки обработки XML с товаром
        /// </summary>
        private void LoadProductXLSButton_Click(object sender, EventArgs e)
        {
            //Запускаем процесс конвертации XML файлов с товарами
            Thread_Dir_Processing = new Thread(Core.Thread_Dir_Processing_Code);
            Thread_Dir_Processing.Start(this);
        }

               



        private void comboBox4_MouseEnter(object sender, EventArgs e)
        {
            comboBox4.Items.Clear();
            comboBox4.Items.Add("ВСЕ");

            foreach(CategoryOfProduct category in mainCategoryList)
            {
                comboBox4.Items.Add(category.Name);
            }

        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();

            if ("ВСЕ" == comboBox2.SelectedItem.ToString())
            {
                foreach (ReplaceKeys RK in Replace_Keys)
                {
                    {
                        //=========== Конструктор для добавления в датагрид =============
                        //Строка для добавления на грид
                        string[] Filtr = new string[4];
                        //Ключ замены регулярное выражение
                        Filtr[0] = RK.RegKey.Value;
                        //Filtr[0] = PL.ChildNodes[0].InnerText;
                        //Значение замены
                        Filtr[1] = RK.NewValue;
                        //Действие
                        Filtr[2] = Stuff.ActionCoder(RK.Action);

                        Filtr[3] = GUIDReplaceKey.ToString();
                        GUIDReplaceKey++;

                        //Добавляем правило в ГРИД
                        dataGridView2.Rows.Add(Filtr);
                        //================================================================
                    }

                }

            }
            else
            {
                List<ReplaceKeys> M = new List<ReplaceKeys>();

                foreach (ReplaceKeys RK in Replace_Keys)
                {
                    if (RK.GroupValue == comboBox2.SelectedItem.ToString())
                    {
                        M.Add(RK);
                    }

                }

                foreach (ReplaceKeys RK in M)
                {
                    //=========== Конструктор для добавления в датагрид =============
                    //Строка для добавления на грид
                    string[] Filtr = new string[4];
                    //Ключ замены регулярное выражение
                    Filtr[0] = RK.RegKey.Value;
                    //Filtr[0] = PL.ChildNodes[0].InnerText;
                    //Значение замены
                    Filtr[1] = RK.NewValue;
                    //Действие
                    Filtr[2] = Stuff.ActionCoder(RK.Action);

                    Filtr[3] = GUIDReplaceKey.ToString();
                    GUIDReplaceKey++;

                    //Добавляем правило в ГРИД
                    dataGridView2.Rows.Add(Filtr);
                    //================================================================

                }
            }
        }

        private void comboBox5_SelectionChangeCommitted(object sender, EventArgs e)
        {
            //if ("ВСЕ" == comboBox5.SelectedItem.ToString())
            //{
            //    listBox3.Items.Clear();

            //    foreach (Product P in ProductListSource)
            //    {
            //        listBox3.Items.Add(P.IDURL);
            //    }
            //}
            //else
            //{
            //    listBox3.Items.Clear();

            //    foreach (Product P in ProductListSource)
            //    {
            //        if (P.IDURL.IndexOf(comboBox5.SelectedItem.ToString())>=0)
            //        {
            //            listBox3.Items.Add(P.IDURL);
            //        }
            //    }
            //}
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void удалитьToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (dataGridView9.SelectedRows.Count > 0)
            {

                int u = dataGridView9.SelectedRows[0].Index;
                int i = u;

                dataGridView9.Rows.RemoveAt(i);


                mainCategoryList[selectedIndexCategory].SubCategoty[selectedIndexSubCategory].Keys.RemoveAt(i);
               // Replace_Keys.RemoveAt(i);
            }

        }

        private void удалитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (dataGridView8.SelectedRows.Count > 0)
            {

                int u = dataGridView8.SelectedRows[0].Index;
                int i = u;

                dataGridView8.Rows.RemoveAt(i);


                mainCategoryList[selectedIndexCategory].SubCategoty.RemoveAt(selectedIndexSubCategory);//[selectedIndexSubCategory].
                // Replace_Keys.RemoveAt(i);
            }

        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (dataGridView6.SelectedRows.Count > 0)
            {

                int u = dataGridView6.SelectedRows[0].Index;
                int i = u;

                dataGridView6.Rows.RemoveAt(i);


                mainCategoryList[selectedIndexCategory].Keys.RemoveAt(i);
                // Replace_Keys.RemoveAt(i);
            }


        }

        //Добавляем новую категорию товара
        private void button16_Click(object sender, EventArgs e)
        {

            if (textBox6.Text.Length >= 3)
            {
                bool categoryExist = true;

                foreach(CategoryOfProduct c in mainCategoryList)
                {
                    if(textBox6.Text == c.Name)
                    {
                        categoryExist = false;
                        MessageBox.Show("Внимание! Категория с таким именем уже существует.");
                        break;
                    }
                }

                if (categoryExist)
                {
                    CategoryOfProduct newCategory = new CategoryOfProduct();
                    newCategory.Name = textBox6.Text;//dataGridView7.Rows[mainCategoryList.Count].Cells[0].Value.ToString();

                    dataGridView7.Rows.Add(new string[] { newCategory.Name, "0" });
                    mainCategoryList.Add(newCategory);
                    catListBox.Items.Add(newCategory.Name);
                }
            }
            else
            {
                MessageBox.Show("Минемальная длина имени категории 3 символа");
            }
        }

        //Удаляем категорию товара
        private void button17_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;

            dataGridView7.Rows.RemoveAt(u);
            mainCategoryList.RemoveAt(u);
            catListBox.Items.RemoveAt(u);
        }
        //Вносим изменения в категорию товара
        private void button18_Click(object sender, EventArgs e)
        {

            int u = dataGridView7.SelectedRows[0].Index;

            bool categoryExist = true;

            foreach (CategoryOfProduct c in mainCategoryList)
            {
                if (textBox6.Text == c.Name)
                {
                    categoryExist = false;
                    MessageBox.Show("Внимание! Категория с таким именем уже существует.");
                    break;
                }
            }

            if (categoryExist)
            {
                dataGridView7.Rows[u].Cells[0].Value = textBox6.Text;
                mainCategoryList[u].Name = textBox6.Text;
                catListBox.Items[u] = textBox6.Text;

            }
        }

        //Добавляем новое регулярное вырожение категорию товара
        private void button20_Click(object sender, EventArgs e)
        {

            int u = dataGridView7.SelectedRows[0].Index;

            bool categoryExist = true;


            foreach (Key k in mainCategoryList[u].Keys)
            {
                if (textBox8.Text == k.Value)
                {
                    categoryExist = false;
                    MessageBox.Show("Внимание! Подобное регулярное выражение уже добавлено.");
                    break;
                }
            }

            if (categoryExist)
            {


                if (textBox8.Text.Length >= 3)
                {
                    Key newKey = new Key();
                    newKey.IsActiv = true;
                    newKey.Value = textBox8.Text;

                    // mainCategoryList[selectedIndexCategory].Keys.Add(newKey);
                    dataGridView6.Rows.Add(new string[] { newKey.Value });
                    mainCategoryList[u].Keys.Add(newKey);
                }
                else
                {
                    MessageBox.Show("Минемальная длина регулярного выражения 3 символа");
                }
            }
        }

        //Удаляем регулярное вырожение категории товаров
        private void button21_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView6.SelectedRows[0].Index;

            mainCategoryList[u].Keys.RemoveAt(w);
            dataGridView6.Rows.RemoveAt(w);
        }

        //Вносим изменения в  регулярное вырожение  категории товара
        private void button22_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView6.SelectedRows[0].Index;

            bool categoryExist = true;


            foreach (Key k in mainCategoryList[u].Keys)
            {
                if (textBox8.Text == k.Value)
                {
                    categoryExist = false;
                    MessageBox.Show("Внимание! Подобное регулярное выражение уже добавлено.");
                    break;
                }
            }

            if (categoryExist)
            {

                dataGridView6.Rows[w].Cells[0].Value = textBox8.Text;
                mainCategoryList[u].Keys[w].Value = textBox8.Text;
            }

        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int w = dataGridView6.SelectedRows[0].Index;

                textBox8.Text = dataGridView6.Rows[w].Cells[0].Value.ToString();
            }
            catch { }
        }

        //Добавляем новую подкатегорию товара
        private void button23_Click(object sender, EventArgs e)
        {
            if (textBox9.Text.Length >= 3)
            {
                bool categoryExist = true;
                int u = dataGridView7.SelectedRows[0].Index;


                foreach (SubCategoryOfProduct sc in mainCategoryList[u].SubCategoty)
                {
                    if (textBox9.Text == sc.Name)
                    {
                        categoryExist = false;
                        MessageBox.Show("Внимание! Подкатегория с таким именем уже существует.");
                        break;
                    }
                }

                if (categoryExist)
                {



                    SubCategoryOfProduct newSubCP = new SubCategoryOfProduct();
                    newSubCP.Name = textBox9.Text;

                    mainCategoryList[u].SubCategoty.Add(newSubCP);
                    dataGridView8.Rows.Add(new string[] { newSubCP.Name, "0" });

                    if (catListBox.SelectedIndex == u)
                    {
                        subCatListBox.Items.Add(newSubCP.Name);
                    }
                }

                // CategoryOfProduct newCategory = new CategoryOfProduct();
                //newCategory.Name = textBox9.Text;//dataGridView7.Rows[mainCategoryList.Count].Cells[0].Value.ToString();

               // mainCategoryList.Add(newCategory);
            }
            else
            {
                MessageBox.Show("Минемальная длина имени категории 3 символа");
            }


          
        }

        private void button25_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView8.SelectedRows[0].Index;


            dataGridView8.Rows.RemoveAt(w);
            mainCategoryList[u].SubCategoty.RemoveAt(w);

            if (catListBox.SelectedIndex == u)
            {
                subCatListBox.Items.RemoveAt(w);
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView8.SelectedRows[0].Index;

            bool categoryExist = true;


            foreach (SubCategoryOfProduct sc in mainCategoryList[u].SubCategoty)
            {
                if (textBox9.Text == sc.Name)
                {
                    categoryExist = false;
                    MessageBox.Show("Внимание! Подкатегория с таким именем уже существует.");
                    break;
                }
            }

            if (categoryExist)
            {


                dataGridView8.Rows[w].Cells[0].Value = textBox9.Text;
                mainCategoryList[u].SubCategoty[w].Name = textBox9.Text;

                if (catListBox.SelectedIndex == u)
                {
                    subCatListBox.Items[w] = textBox9.Text;
                }

                //try
                //{
                //    if (dataGridView8.Rows.Count == mainCategoryList[selectedIndexCategory].SubCategoty.Count + 1)
                //    {
                //        int i = 0;
                //        foreach (DataGridViewRow r in dataGridView8.Rows)
                //        {
                //            mainCategoryList[selectedIndexCategory].SubCategoty[i].Name = r.Cells[0].Value.ToString();

                //            i++;
                //        }
                //    }

                //}
                //catch { }
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView8.SelectedRows[0].Index;
            //int f = dataGridView9.SelectedRows[0].Index;

            bool categoryExist = true;


            foreach (Key k in mainCategoryList[u].SubCategoty[w].Keys)
            {
                if (textBox10.Text == k.Value)
                {
                    categoryExist = false;
                    MessageBox.Show("Внимание! Подобное регулярное выражение уже добавлено.");
                    break;
                }
            }

            if (categoryExist)
            {

                if (textBox10.Text.Length >= 3)
                {
                    Key newKey = new Key();
                    newKey.IsActiv = true;
                    newKey.Value = textBox10.Text;

                    // mainCategoryList[selectedIndexCategory].Keys.Add(newKey);
                    dataGridView9.Rows.Add(new string[] { newKey.Value });
                    mainCategoryList[u].SubCategoty[w].Keys.Add(newKey);
                }
                else
                {
                    MessageBox.Show("Минемальная длина регулярного выражения 3 символа");
                }
            }

        }

        private void button26_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView8.SelectedRows[0].Index;
            int f = dataGridView9.SelectedRows[0].Index;

            mainCategoryList[u].SubCategoty[w].Keys.RemoveAt(f);
            dataGridView9.Rows.RemoveAt(f);

        }

        private void button28_Click(object sender, EventArgs e)
        {
            int u = dataGridView7.SelectedRows[0].Index;
            int w = dataGridView8.SelectedRows[0].Index;
            int f = dataGridView9.SelectedRows[0].Index;

            bool categoryExist = true;


            foreach (Key k in mainCategoryList[u].SubCategoty[w].Keys)
            {
                if (textBox10.Text == k.Value)
                {
                    categoryExist = false;
                    MessageBox.Show("Внимание! Подобное регулярное выражение уже добавлено.");
                    break;
                }
            }

            if (categoryExist)
            {

                dataGridView9.Rows[f].Cells[0].Value = textBox10.Text;
                mainCategoryList[u].SubCategoty[w].Keys[f].Value = textBox10.Text;
            }
        }

        private void dataGridView9_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int f = dataGridView9.SelectedRows[0].Index;
                textBox10.Text = dataGridView9.Rows[f].Cells[0].Value.ToString();
            }
            catch { }
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            int indexOfProduct = -1;

            int j = 0;
            foreach(Product p in ProductListForPosting)
            {
                try
                {
                    foreach (string photopath in p.FilePath)
                    {
                        if (photopath == listView2.SelectedItems[0].Text)
                        {
                            indexOfProduct = j;
                            break;
                        }
                    }
                    j++;

                    if (indexOfProduct >= 0)
                    {
                        break;
                    }
                }
                catch { }
            }


            if (indexOfProduct>=0)
            {
                textBox7.Text = "";
                textBox5.Text = "";
                listBox5.Items.Clear();
                listBox4.Items.Clear();
                int ic = 0;

                int indexcat = -1;
                int indexSUB = -1;

                try
                {
                    foreach (CategoryOfProduct cat in mainCategoryList)
                    {
                        listBox5.Items.Add(cat.Name);

                        if (cat.Name == ProductListForPosting[indexOfProduct].CategoryOfProductName)
                        {
                            indexcat = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }
                ic = 0;
                try
                {
                    foreach (SubCategoryOfProduct SUB in mainCategoryList[indexcat].SubCategoty)
                    {
                        listBox4.Items.Add(SUB.Name);

                        if (SUB.Name == ProductListForPosting[indexOfProduct].SubCategoryOfProductName)
                        {
                            indexSUB = ic;
                        }
                        ic++;
                    }
                }
                catch
                { }


                listBox5.SelectedIndex = indexcat;
                listBox4.SelectedIndex = indexSUB;



                ////Чистим картинки в листбоксе
                //imageList1.Images.Clear();
                //listView2.Items.Clear();

                //int ii = 0;
                //foreach (string s in ProductListForPosting[indexOfProduct].FilePath)
                //{
                //    imageList1.Images.Add(new Bitmap(s));
                //    listView2.Items.Add(new ListViewItem(s, ii));
                //    ii++;
                //}

                string sout = "";
                foreach (string sss in ProductListForPosting[indexOfProduct].sellerTextCleen)
                {
                    sout = sout + sss + "\r\n";
                }
                textBox7.Text = sout;


                sout = "";
                foreach (string sss in ProductListForPosting[indexOfProduct].sellerText)
                {
                    sout = sout + sss + "\r\n";
                }
                textBox11.Text = sout;



                textBox5.Text = ProductListForPosting[indexOfProduct].Prises;

                imageList2.Images.Clear();
                listView3.Items.Clear();

                int ii = 0;
                foreach (string s in ProductListForPosting[indexOfProduct].FilePath)
                {
                    imageList2.Images.Add(new Bitmap(s));
                    listView3.Items.Add(new ListViewItem(s, ii));
                    ii++;
                }

            }

        }

        private void listView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {

                int indexOfProduct = -1;

                int j = 0;
                foreach (Product p in ProductListForPosting)
                {
                    try
                    {
                        foreach (string photopath in p.FilePath)
                        {
                            if (photopath == listView2.SelectedItems[0].Text)
                            {
                                indexOfProduct = j;
                                break;
                            }
                        }
                        j++;

                        if (indexOfProduct >= 0)
                        {
                            break;
                        }
                    }
                    catch { }
                }


                if (indexOfProduct >= 0)
                {

                      
                        
                        
                    ProductListForPosting[indexOfProduct].URLPhoto.RemoveAt(listView3.SelectedIndices[0]);
                    ProductListForPosting[indexOfProduct].FilePath.RemoveAt(listView3.SelectedIndices[0]);


                    listView3.Items.Remove(listView3.SelectedItems[0]);
                    listView2.Items.Remove(listView2.SelectedItems[0]);

                }







            }

        }

        private void button31_Click(object sender, EventArgs e)
        {
            if ((dataGridView7.SelectedRows.Count > 0) && (dataGridView7.Rows.Count == mainCategoryList.Count))
            {
                int u = dataGridView7.SelectedRows[0].Index;
                int i = u;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                if (i > 0)
                {

                    string[] s = new string[2];
                    s[0] = dataGridView7.Rows[i].Cells[0].Value.ToString();
                    s[1] = dataGridView7.Rows[i].Cells[1].Value.ToString();

                    dataGridView7.Rows.Insert(i - 1, s);
                    dataGridView7.Rows.RemoveAt(i + 1);
                    // dataGridView2.ClearSelection();

                    //i = 0;
                    ////int j = 0;
                    //foreach (ReplaceKeys r in Replace_Keys)
                    //{
                    //    if (r.RegKey.Value == dataGridView2.Rows[u].Cells[0].Value.ToString())
                    //    {
                    //        break;
                    //    }
                    //    i++;
                    //}
                    //int i = selectedIndexReplace;

                    mainCategoryList.Insert(i - 1, mainCategoryList[i]);
                    mainCategoryList.RemoveAt(i + 1);

                    dataGridView7.ClearSelection();

                    if (u - 1 >= 0)
                    {
                        dataGridView7.Rows[u - 1].Selected = true;
                    }
                    else
                    {

                        dataGridView7.Rows[0].Selected = true;

                    }
                }
            }


            //listBox6.Items.Clear();
            //foreach (ReplaceKeys r in Replace_Keys)
            //{
            //    listBox6.Items.Add(r.RegKey.Value);
            //}

            UpdateMainCategoryList();

        }




        public void UpdateMainCategoryList()
        {
            catListBox.Items.Clear();
            subCatListBox.Items.Clear();
            foreach (CategoryOfProduct c in mainCategoryList)
            {
                catListBox.Items.Add(c.Name);
            }
        }




        private void button32_Click(object sender, EventArgs e)
        {
            if ((dataGridView7.SelectedRows.Count > 0) && (dataGridView7.Rows.Count == mainCategoryList.Count))
            {
                int u = dataGridView7.SelectedRows[0].Index;
                int i = u;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                if (i < dataGridView7.Rows.Count - 1)
                {

                    string[] s = new string[2];
                    s[0] = dataGridView7.Rows[i].Cells[0].Value.ToString();
                    s[1] = dataGridView7.Rows[i].Cells[1].Value.ToString();

                    dataGridView7.Rows.Insert(i + 2, s);
                    dataGridView7.Rows.RemoveAt(i);
                    // dataGridView2.ClearSelection();

                    //i = 0;
                    ////int j = 0;
                    //foreach (ReplaceKeys r in Replace_Keys)
                    //{
                    //    if (r.RegKey.Value == dataGridView2.Rows[u].Cells[0].Value.ToString())
                    //    {
                    //        break;
                    //    }
                    //    i++;
                    //}
                    //int i = selectedIndexReplace;

                    mainCategoryList.Insert(i + 2, mainCategoryList[i]);
                    mainCategoryList.RemoveAt(i);


                    dataGridView7.ClearSelection();

                    if (u + 1 < dataGridView7.Rows.Count)
                    {
                        dataGridView7.Rows[u + 1].Selected = true;
                    }
                    else
                    {
                        dataGridView7.Rows[dataGridView7.Rows.Count - 1].Selected = true;
                    }
                }
            }
            //listBox6.Items.Clear();
            //foreach (ReplaceKeys r in Replace_Keys)
            //{
            //    listBox6.Items.Add(r.RegKey.Value);
            //}

            UpdateMainCategoryList();

        }

        private void button34_Click(object sender, EventArgs e)
        {
            if ((dataGridView8.SelectedRows.Count > 0) && (dataGridView7.SelectedRows.Count > 0))
            {
                int u = dataGridView7.SelectedRows[0].Index;
                int w = dataGridView8.SelectedRows[0].Index;
                int i = w;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                if (i > 0)
                {

                    string[] s = new string[1];
                    s[0] = dataGridView8.Rows[i].Cells[0].Value.ToString();

                    dataGridView8.Rows.Insert(i - 1, s);
                    dataGridView8.Rows.RemoveAt(i + 1);
                    // dataGridView2.ClearSelection();

                    //i = 0;
                    ////int j = 0;
                    //foreach (ReplaceKeys r in Replace_Keys)
                    //{
                    //    if (r.RegKey.Value == dataGridView2.Rows[u].Cells[0].Value.ToString())
                    //    {
                    //        break;
                    //    }
                    //    i++;
                    //}
                    //int i = selectedIndexReplace;

                    mainCategoryList[u].SubCategoty.Insert(i - 1, mainCategoryList[u].SubCategoty[w]);
                    mainCategoryList[u].SubCategoty.RemoveAt(i + 1);

                    dataGridView8.ClearSelection();

                    if (w - 1 >= 0)
                    {
                        dataGridView8.Rows[w - 1].Selected = true;
                    }
                    else
                    {

                        dataGridView8.Rows[0].Selected = true;

                    }
                }
            }


            //listBox6.Items.Clear();
            //foreach (ReplaceKeys r in Replace_Keys)
            //{
            //    listBox6.Items.Add(r.RegKey.Value);
            //}

            UpdateMainCategoryList();

        }

        private void button33_Click(object sender, EventArgs e)
        {
            if ((dataGridView8.SelectedRows.Count > 0) && (dataGridView7.SelectedRows.Count > 0))
            {
                int u = dataGridView7.SelectedRows[0].Index;
                int w = dataGridView8.SelectedRows[0].Index;
                int i = w;
                //int j = 0;
                //foreach (DataGridViewRow r in dataGridView2.Rows)
                //{
                //    if (r.Cells[0].Value == dataGridView2.SelectedRows[0].Cells[0].Value)
                //    {
                //        break;
                //    }
                //    i++;
                //}
                //int i = selectedIndexReplace;

                if (i < dataGridView8.Rows.Count - 1)
                {

                    string[] s = new string[1];
                    s[0] = dataGridView8.Rows[i].Cells[0].Value.ToString();

                    dataGridView8.Rows.Insert(i + 2, s);
                    dataGridView8.Rows.RemoveAt(i);
                    // dataGridView2.ClearSelection();

                    //i = 0;
                    ////int j = 0;
                    //foreach (ReplaceKeys r in Replace_Keys)
                    //{
                    //    if (r.RegKey.Value == dataGridView2.Rows[u].Cells[0].Value.ToString())
                    //    {
                    //        break;
                    //    }
                    //    i++;
                    //}
                    //int i = selectedIndexReplace;

                    mainCategoryList[u].SubCategoty.Insert(i + 2, mainCategoryList[u].SubCategoty[i]);
                    mainCategoryList[u].SubCategoty.RemoveAt(i);


                    dataGridView8.ClearSelection();

                    if (w + 1 < dataGridView8.Rows.Count)
                    {
                        dataGridView8.Rows[w + 1].Selected = true;
                    }
                    else
                    {
                        dataGridView8.Rows[dataGridView8.Rows.Count - 1].Selected = true;
                    }
                }
            }
            //listBox6.Items.Clear();
            //foreach (ReplaceKeys r in Replace_Keys)
            //{
            //    listBox6.Items.Add(r.RegKey.Value);
            //}

            UpdateMainCategoryList();

        }


        /// <summary>
        /// Метод загрузки данных из XML файла с поставщиками
        /// </summary>
        private void LoadProviderXLSButton_Click(object sender, EventArgs e)
        {
            //Запускаем поток конвертации файла с постовщиками во внутренний формат
            Thread_Provider_Processing = new Thread(Thread_Provider_Processing_Code);
            Thread_Provider_Processing.Start();
        }

        public void Thread_Provider_Processing_Code()
        {

            Stuff.ExportProviderExcel(this);








            //Action S1 = () => button1.Enabled = false;
            //button1.Invoke(S1);

            //DirectoryInfo d = new DirectoryInfo(_InputPath);

            //if (!d.Exists)
            //{
            //    MessageBox.Show("Директория " + _InputPath + " не существует!");
            //}
            //else
            //{

            //    int cf = d.GetFiles().Length;
            //    int pf = 0;

            //    imageNoExist.Clear();


            //    foreach (FileInfo f in d.GetFiles())
            //    {
            //        try
            //        {
            //            Action S2 = () => label11.Text = "Файлов обработано " + pf.ToString() + " из " + cf.ToString();
            //            label11.Invoke(S2);




            //           Stuff.ExportExcel(f.FullName,this);



            //            f.Delete();
            //            pf++;
            //        }
            //        catch
            //        { }
            //    }


            //    foreach (Product p in ProductListSourceBuffer)
            //    {
            //        ProductListSource.Add(p);
            //    }


            //    ProductListSourceBuffer.Clear();


            //    string sL = "При загрузке отсутствуют следующие изображения: \r\n";
            //    foreach (string L in imageNoExist)
            //    {
            //        sL = sL + L + "\r\n";
            //    }

            //    MessageBox.Show(sL);

            //    listBox3.Items.Clear();

            //    int i = 1;

            //    foreach (Product P in ProductListSource)
            //    {
            //        listBox3.Items.Add(i);
            //        i++;
            //    }

            //    Stuff.UpdatePostavshikov(this);

            //    Action S3 = () => button1.Enabled = true;
            //    button1.Invoke(S3);
            //}

        }

    }
}

