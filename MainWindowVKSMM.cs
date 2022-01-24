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

namespace VKSMM
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ExportExcel();
        }


        public class Product
        {
            public string IDURL = string.Empty;
            public List<string> URLPhoto = new List<string>();
            public List<string> FilePath = new List<string>();
            public DateTime datePost;
            public int[] prise = new int[5];
            public List<string> sellerText = new List<string>();
            public List<string> sellerTextCleen = new List<string>();
            public List<string> logRegularExpression = new List<string>();
            public string CategoryOfProductName = "ВСЕ";
            public string SubCategoryOfProductName = "ВСЕ";
            public bool HandBlock = false;

            public string Materials = "";
            public string Prises = "";
            public string Sizes = "";
        }


        public class CategoryOfProduct
        {
            public CategoryOfProduct()
            {
                SubCategoty.Add(new SubCategoryOfProduct());
            }

            public bool isProvider = true;

            public string Name = "Всякое";
            //public string NamespaceHandling = "Всякое";
            public List<SubCategoryOfProduct> SubCategoty = new List<SubCategoryOfProduct>();
            public List<Key> Keys = new List<Key>();
        }

        public class SubCategoryOfProduct
        {
            public string Name = "ВСЕ";
            public List<Key> Keys = new List<Key>();
        }

        public class Key
        {
            public bool isProvider = true;
            public bool IsActiv = true;
            public string Value = "";
        }

        public class ColorKeys
        {
            public string Name = "";
            public Color color = Color.White;
            public int Action = -1;
            public Key RegKey = new Key();
        }

        public class ReplaceKeys
        {
            public string Name = "";
            public string NewValue = "";
            public string GroupValue = "ВСЕ";
            public int Action = -1;
            public Key RegKey = new Key();
        }


        public Thread Thread_Dir_Processing; 
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

        public List<Product> ProductListSource = new List<Product>();

        public List<Product> ProductListSourceBuffer = new List<Product>();
        

        public List<Product> ProductListForPosting = new List<Product>();

        public List<ReplaceKeys> Replace_Keys = new List<ReplaceKeys>();

        //public List<ReplaceKeys> Addition_Replace_Keys = new List<ReplaceKeys>();

        public List<ColorKeys> Color_Keys = new List<ColorKeys>();

        //Декодер действия
        public int ActionDecoder(string CodeAct)
        {
            if (CodeAct == "Блокировать") return 5;
            if (CodeAct == "Заменять") return 3;
            if (CodeAct == "Дописывать") return 4;
            if (CodeAct == "Пропускать") return 2;
            if (CodeAct == "Удалять") return 1;
            return -1;
        }

        //Кодер действия
        public string ActionCoder(int CodeAct)
        {
            if (CodeAct == 5) return "Блокировать";
            if (CodeAct == 3) return "Заменять";
            if (CodeAct == 4) return "Дописывать";
            if (CodeAct == 2) return "Пропускать";
            if (CodeAct == 1) return "Удалять";
            return "Пусто";
        }

        public List<string> ConvertMassToList(string[] Lines)
        {
            List<string> OutLine = new List<string>();

            foreach (string S in Lines)
            {
                OutLine.Add(S);
            }

            return OutLine;
        }

        public int[] ConvertMassToInt(string[] Lines)
        {
            int[] OutLine = new int[Lines.Length];

            for (int i = 0; i < Lines.Length; i++)
            {
                OutLine[i] = Convert.ToInt32(Lines[i]);
            }

            return OutLine;
        }


        private void MainForm_Load(object sender, EventArgs e)
        {


            if (DateTime.Now.Year > 2021 && DateTime.Now.Month > 11)
            {
                MessageBox.Show("Ошибка! Обратитесь к разработчику!");
                this.Close();
            }


            ////Сохраняем исходное имя формы
            ////_IshodFormName = this.Text;

            //CategoryOfProduct c1 = new CategoryOfProduct();
            //c1.Name = "Женское";
            //CategoryOfProduct c2 = new CategoryOfProduct();
            //c2.Name = "Мужское";

            //Key k1 = new Key(); k1.Value = "юбк[аио]";
            //c1.Keys.Add(k1);

            //Key k2 = new Key(); k2.Value = "женск";
            //c1.Keys.Add(k2);

            //SubCategoryOfProduct s1 = new SubCategoryOfProduct();
            //s1.Name = "Сумки";

            //Key sk1 = new Key(); sk1.Value = "сумк";
            //s1.Keys.Add(sk1);

            //SubCategoryOfProduct s2 = new SubCategoryOfProduct();
            //s2.Name = "Аксесуары";

            //c1.SubCategoty.Add(s1);
            //c1.SubCategoty.Add(s2);
            ////Key sk1 = new Key(); sk1.Value = "\\<сумк";
            ////s1.Keys.Add(sk1);


            //mainCategoryList.Add(c1);
            //mainCategoryList.Add(c2);

            //listBox1.Items.Add(c1.Name);
            //listBox1.Items.Add(c2.Name);


            //XML документ с настройками программы
            XmlDocument conf_supp = new XmlDocument();
            try
            {
                //Создаем ссылку на файл настроек
                FileInfo F = new FileInfo("ConfigTKSadovod.XML");


                DirectoryInfo d = new DirectoryInfo("CopyOfConfig");
                if (!d.Exists) d.Create();
                F.CopyTo(d.FullName+"\\"+DateTime.Now.Date.ToShortDateString().Replace(".","")+ DateTime.Now.TimeOfDay.ToString().Replace(":","").Replace(".","")+".xml");



                //Сохраняем путь к файлу настроек
                _path = F.FullName;
                //Читаем файл настроек в память
                conf_supp.Load("ConfigTKSadovod.XML");

                //=======================================================================================================
                XmlNodeList nodeList;//Вспомогательнаяпеременная
                XmlNode root = conf_supp.DocumentElement;//Получаем доступ к XML файлу настроек


                _ProductDBPath = F.FullName.Substring(0, F.FullName.LastIndexOf("\\"));

                try
                {
                    nodeList = root.SelectNodes("PRODUCT_DB_DIR");//Считываем настройки ключей замены

                    _ProductDBPath = nodeList[0].InnerText;
                }
                catch
                {

                }

                try
                {
                    nodeList = root.SelectNodes("PHOTO_DIR");//Считываем настройки ключей замены

                    _PhotoPath = nodeList[0].InnerText;
                }
                catch
                {

                }
                try
                {
                    nodeList = root.SelectNodes("INPUT_DIR");//Считываем настройки ключей замены

                    _InputPath = nodeList[0].InnerText;
                }
                catch
                {

                }

                try
                {
                    nodeList = root.SelectNodes("PROVIDER_DIR");//Считываем настройки ключей замены

                    _ProviderDir = nodeList[0].InnerText;
                }
                catch
                {

                }


                try
                {
                    nodeList = root.SelectNodes("REPLACE_KEYS");//Считываем настройки ключей замены
                    //Пробегаем по всем ключам замены
                    foreach (XmlNode PL in nodeList[0].ChildNodes)
                    {
                        //=========== Конструктор для добавления в датагрид =============
                        //Строка для добавления на грид
                        string[] Filtr = new string[5];
                        //Ключ замены регулярное выражение
                        Filtr[0] = Encoding.Unicode.GetString(Convert.FromBase64String(PL.ChildNodes[0].InnerText));
                        //Filtr[0] = PL.ChildNodes[0].InnerText;
                        //Значение замены
                        Filtr[1] = PL.ChildNodes[1].InnerText;
                        //Действие
                        Filtr[2] = PL.ChildNodes[2].InnerText;

                        Filtr[3] = GUIDReplaceKey.ToString();

                        Filtr[4] = PL.ChildNodes[3].InnerText;

                        GUIDReplaceKey++;

                        //Добавляем правило в ГРИД
                        dataGridView2.Rows.Add(Filtr);

                        //================================================================

                        //=========== Конструктор для добавления в датагрид =============
                        //Создаем экземпляр ключа
                        Key k = new Key();
                        //Значение ключа
                        k.Value = Filtr[0].Replace("\r","").Replace("\n","");
                        //Флаг активности ключа
                        k.IsActiv = true;
                        
                        //Класс замены
                        ReplaceKeys r = new ReplaceKeys();
                        //Действие связанное с заменой 3-ка просто замена
                        r.Action = ActionDecoder(Filtr[2]);
                        //Ключ приыязанный к классу замены
                        r.RegKey = k;
                        //Значение замены
                        r.NewValue = Filtr[1].Replace("\r", "").Replace("\n", "");

                        try
                        {
                            r.GroupValue = PL.ChildNodes[3].InnerText;
                        }
                        catch { }

                        AddGroup(r);

                        //Добавляем ключ в пул замен
                        Replace_Keys.Add(r);
                        //================================================================
                    }
                }
                catch { }

                try
                {
                    nodeList = root.SelectNodes("COLOR_KEYS");//Считываем настройки цветовых ключей
                    //Пробегаем по всем ключам цветовой дифференциации
                    foreach (XmlNode PL in nodeList[0].ChildNodes)
                    {

                        //=========== Конструктор для добавления в датагрид =============
                        //Строка для добавления на грид
                        string[] Filtr = new string[2];
                        //Ключ замены регулярное выражение
                        Filtr[0] = Encoding.Unicode.GetString(Convert.FromBase64String(PL.ChildNodes[0].InnerText));
                        //Filtr[0] = PL.ChildNodes[0].InnerText;
                        //Действие
                        Filtr[1] = PL.ChildNodes[1].InnerText;
                        //Добавляем правило в ГРИД
                        dataGridView4.Rows.Add(Filtr);
                        //Окрашываем строчку
                        dataGridView4.Rows[dataGridView4.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromName(PL.ChildNodes[2].InnerText);
                        //================================================================

                        //=========== Конструктор для добавления в датагрид =============
                        //Создаем экземпляр ключа
                        Key k = new Key();
                        //Значение ключа
                        k.Value = Filtr[0];
                        //Флаг активности ключа
                        k.IsActiv = true;

                        //Класс замены
                        ColorKeys r = new ColorKeys();
                        //Действие связанное с заменой 3-ка просто замена
                        r.Action = ActionDecoder(Filtr[1]);
                        //Ключ приыязанный к классу замены
                        r.RegKey = k;
                        //Значение замены
                        r.color = Color.FromName(PL.ChildNodes[2].InnerText);

                        //Добавляем ключ в пул замен
                        Color_Keys.Add(r);
                        //================================================================
                    }
                }
                catch { }



                //try
                //{
                //    nodeList = root.SelectNodes("ADDITION_KEYS");//Считываем настройки ключей замены
                //    //Пробегаем по всем ключам замены
                //    foreach (XmlNode PL in nodeList[0].ChildNodes)
                //    {
                //        //=========== Конструктор для добавления в датагрид =============
                //        //Строка для добавления на грид
                //        string[] Filtr = new string[3];
                //        //Ключ замены регулярное выражение
                //        Filtr[0] = Encoding.Unicode.GetString(Convert.FromBase64String(PL.ChildNodes[0].InnerText));
                //        //Filtr[0] = PL.ChildNodes[0].InnerText;
                //        //Значение замены
                //        Filtr[1] = PL.ChildNodes[1].InnerText;
                //        //Действие
                //        Filtr[2] = PL.ChildNodes[2].InnerText;
                //        //Добавляем правило в ГРИД
                //        dataGridView4.Rows.Add(Filtr);
                //        //================================================================

                //        //=========== Конструктор для добавления в датагрид =============
                //        //Создаем экземпляр ключа
                //        Key k = new Key();
                //        //Значение ключа
                //        k.Value = Filtr[0];
                //        //Флаг активности ключа
                //        k.IsActiv = true;

                //        //Класс замены
                //        ReplaceKeys r = new ReplaceKeys();
                //        //Действие связанное с заменой 3-ка просто замена
                //        r.Action = ActionDecoder(Filtr[2]);
                //        //Ключ приыязанный к классу замены
                //        r.RegKey = k;
                //        //Значение замены
                //        r.NewValue = Filtr[1];

                //        //Добавляем ключ в пул замен
                //        Addition_Replace_Keys.Add(r);
                //        //================================================================
                //    }
                //}
                //catch { }


                try
                {
                    nodeList = root.SelectNodes("CATEGORY_PRODUCT");//Считываем категории продуктов
                    //Пробегаем по всем категориям продуктов
                    foreach (XmlNode PL in nodeList[0].ChildNodes)
                    {

                        CategoryOfProduct COFP = new CategoryOfProduct();
                        COFP.Name = PL.ChildNodes[0].InnerText;

                        try
                        {
                            if (PL.ChildNodes[1].ChildNodes.Count > 0)
                            {
                                foreach (XmlNode KEYPL in PL.ChildNodes[1].ChildNodes)
                                {
                                    //Создаем экземпляр ключа
                                    Key kmc = new Key();
                                    //Значение ключа
                                    kmc.Value = Encoding.Unicode.GetString(Convert.FromBase64String(KEYPL.ChildNodes[0].InnerText));
                                    //kmc.Value = KEYPL.ChildNodes[0].InnerText;
                                    //Флаг активности ключа
                                    kmc.IsActiv = true;

                                    COFP.Keys.Add(kmc);
                                }
                            }
                        }
                        catch { }
                        try
                        {
                            if (PL.ChildNodes[2].ChildNodes.Count > 0)
                            {
                                foreach (XmlNode SUBPL in PL.ChildNodes[2].ChildNodes)
                                {
                                    SubCategoryOfProduct SOFC = new SubCategoryOfProduct();
                                    SOFC.Name = SUBPL.ChildNodes[0].InnerText;

                                    try
                                    {
                                        if (SUBPL.ChildNodes[1].ChildNodes.Count > 0)
                                        {
                                            foreach (XmlNode KEYPL in SUBPL.ChildNodes[1].ChildNodes)
                                            {
                                                //Создаем экземпляр ключа
                                                Key kmc = new Key();
                                                //Значение ключа
                                                kmc.Value = Encoding.Unicode.GetString(Convert.FromBase64String(KEYPL.ChildNodes[0].InnerText));
                                                //kmc.Value = KEYPL.ChildNodes[0].InnerText;
                                                //Флаг активности ключа
                                                kmc.IsActiv = true;

                                                SOFC.Keys.Add(kmc);
                                            }
                                        }
                                    }
                                    catch { }

                                    COFP.SubCategoty.Add(SOFC);
                                }
                            }
                        }
                        catch { }


                        listBox1.Items.Add(COFP.Name);

                        mainCategoryList.Add(COFP);

                        string[] s = new string[2];

                        s[0] = COFP.Name;
                        s[1] = COFP.SubCategoty.Count.ToString();
                        dataGridView7.Rows.Add(s);
                    }


                }
                catch { }

                //=======================================================================================================
                //Загружаем товары из базы данных
                //=======================================================================================================

                FileInfo f = new FileInfo(_ProductDBPath+"\\ProductDB.csv");
                FileStream fileStream = new FileStream(f.FullName,FileMode.Open);
                StreamReader sr = new StreamReader(fileStream,Encoding.UTF8);

                string Line = "";


                int _it = 0; 

                while (!sr.EndOfStream)
                {
                    Line = sr.ReadLine(); try
                    {
                        string[] MassLine = Line.Split(new char[] { '\t' });

                        Product PN = new Product();

                        PN.CategoryOfProductName = MassLine[0];
                        PN.SubCategoryOfProductName = MassLine[1];

                        PN.datePost = Convert.ToDateTime(MassLine[2]);
                        PN.HandBlock = Convert.ToBoolean(MassLine[3]);

                        PN.IDURL = MassLine[4];
                        PN.Materials = MassLine[5];

                        PN.Prises = MassLine[6];
                        PN.Sizes = MassLine[7];


                        PN.prise = ConvertMassToInt(MassLine[8].Split(new char[] { ',' }));

                        PN.sellerText = ConvertMassToList(MassLine[9].Split(new char[] { ',' }));

                        PN.sellerTextCleen = ConvertMassToList(MassLine[10].Split(new char[] { ',' }));

                        PN.URLPhoto = ConvertMassToList(MassLine[11].Split(new char[] { ',' }));

                        PN.FilePath = ConvertMassToList(MassLine[12].Split(new char[] { ',' }));

                        ProductListForPosting.Add(PN);

                        AddToTreeView(PN, _it);
                        _it++;
                    }
                    catch
                    {

                    }
                }
               
                sr.Close();
                fileStream.Close();

                //=======================================================================================================
                

                //=======================================================================================================
                //Загружаем товары из базы данных
                //=======================================================================================================

                FileInfo f_UP = new FileInfo(_ProductDBPath + "\\ProductDBUnProcessed.csv");
                FileStream fileStream_UP = new FileStream(f_UP.FullName, FileMode.Open);
                StreamReader sr_UP = new StreamReader(fileStream_UP, Encoding.UTF8);

                Line = "";


                _it = 0;

                while (!sr_UP.EndOfStream)
                {
                    Line = sr_UP.ReadLine(); try
                    {
                        string[] MassLine = Line.Split(new char[] { '\t' });

                        Product PN = new Product();

                        PN.CategoryOfProductName = MassLine[0];
                        PN.SubCategoryOfProductName = MassLine[1];

                        PN.datePost = Convert.ToDateTime(MassLine[2]);
                        PN.HandBlock = Convert.ToBoolean(MassLine[3]);

                        PN.IDURL = MassLine[4];
                        PN.Materials = MassLine[5];

                        PN.Prises = MassLine[6];
                        PN.Sizes = MassLine[7];


                        PN.prise = ConvertMassToInt(MassLine[8].Split(new char[] { ',' }));

                        PN.sellerText = ConvertMassToList(MassLine[9].Split(new char[] { ',' }));

                        PN.sellerTextCleen = ConvertMassToList(MassLine[10].Split(new char[] { ',' }));

                        PN.URLPhoto = ConvertMassToList(MassLine[11].Split(new char[] { ',' }));

                        PN.FilePath = ConvertMassToList(MassLine[12].Split(new char[] { ',' }));

                        ProductListSource.Add(PN);

                        listBox3.Items.Add(PN.IDURL);
                        _it++;
                    }
                    catch
                    {

                    }
                }

                sr_UP.Close();
                fileStream_UP.Close();

                //=======================================================================================================

            }
            catch//Если при загрузке конфиг файла случился конфуз, сообщаем пользователю и загружаем по умолчанию
            {
                //MessageBox.Show("При загрузке конфигурационного файла произошла ошибка! Загружены настройки по умолчанию! GUID будет сброшен!");
                //textBox_InputPath.Text = "Введите входную директорию!";//Значение по умолчанию
                //textBox_OutputPath.Text = "Введите выходную директорию!";//Значение по умолчанию
                //ButtonCompression.BackColor = Color.LightGray;//Значение по умолчанию
                //ButtonDeCompression.BackColor = Color.LightGray;//Значение по умолчанию
            }

            UpdatePostavshikov();



            ExportProviderExcel();
        }

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
                XmlWriter writer1 = XmlWriter.Create(_path, settings1);
                //Записываем заглавие документа
                writer1.WriteStartDocument();
                //Корневой тег
                writer1.WriteStartElement("CFG");
                //Сохраняем номер копа
                //writer1.WriteElementString("AUTHENTIFICATION_STATUS", _AUTHENTIFICATION_STATUS.ToString());



                //Сохраняем максимальный объем файлов в контейнере
                writer1.WriteElementString("PRODUCT_DB_DIR", _ProductDBPath);

                //Сохраняем максимальный объем файлов в контейнере
                writer1.WriteElementString("PHOTO_DIR", _PhotoPath);

                //Сохраняем максимальный объем файлов в контейнере
                writer1.WriteElementString("INPUT_DIR", _InputPath);

                //Сохраняем максимальный объем файлов в контейнере
                writer1.WriteElementString("PROVIDER_DIR", _ProviderDir);

                //Корневой тег
                writer1.WriteStartElement("REPLACE_KEYS");

                foreach (ReplaceKeys RK in Replace_Keys)
                {
                    //Корневой тег
                    writer1.WriteStartElement("R_KEY");

                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(RK.RegKey.Value)));
                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("VALUE", RK.NewValue);
                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("MODE", ActionCoder(RK.Action));
                    //Сохраняем максимальный объем файлов в контейнере
                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("GROUP", RK.GroupValue);

                    //Закрываем корневой тег
                    writer1.WriteEndElement();
                }
                writer1.WriteEndElement();

                //Корневой тег
                //writer1.WriteStartElement("ADDITION_KEYS");

                //foreach (ReplaceKeys RK in Addition_Replace_Keys)
                //{
                //    //Корневой тег
                //    writer1.WriteStartElement("R_KEY");

                //    //Сохраняем максимальный объем файлов в контейнере
                //    writer1.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(RK.RegKey.Value)));
                //    //Сохраняем максимальный объем файлов в контейнере
                //    writer1.WriteElementString("VALUE", RK.NewValue);
                //    //Сохраняем максимальный объем файлов в контейнере
                //    writer1.WriteElementString("MODE", ActionCoder(RK.Action));
                //    //Сохраняем максимальный объем файлов в контейнере

                //    //Закрываем корневой тег
                //    writer1.WriteEndElement();
                //}
                //writer1.WriteEndElement();



                writer1.WriteStartElement("COLOR_KEYS");

                foreach (ColorKeys CK in Color_Keys)
                {
                    //Корневой тег
                    writer1.WriteStartElement("C_KEY");

                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(CK.RegKey.Value)));
                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("MODE", ActionCoder(CK.Action));
                    //Сохраняем максимальный объем файлов в контейнере
                    writer1.WriteElementString("COLOR", CK.color.Name);
                    //Сохраняем максимальный объем файлов в контейнере

                    //Закрываем корневой тег
                    writer1.WriteEndElement();
                }

                writer1.WriteEndElement();

                writer1.WriteStartElement("CATEGORY_PRODUCT");

                foreach (CategoryOfProduct CP in mainCategoryList)
                {
                    if (CP.isProvider)
                    {

                        //Корневой тег
                        writer1.WriteStartElement("CATEG_KEYS");

                        //Сохраняем максимальный объем файлов в контейнере
                        writer1.WriteElementString("NAME", CP.Name);

                        writer1.WriteStartElement("KEYS");

                        foreach (Key CP_K in CP.Keys)
                        {
                            if (CP_K.isProvider)
                            {

                                writer1.WriteStartElement("CAT_KEY");

                                //Сохраняем максимальный объем файлов в контейнере
                                writer1.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(CP_K.Value)));

                                //Закрываем корневой тег
                                writer1.WriteEndElement();
                            }
                        }

                        //Закрываем корневой тег
                        writer1.WriteEndElement();


                        writer1.WriteStartElement("SUB_KATS");

                        foreach (SubCategoryOfProduct CP_S in CP.SubCategoty)
                        {
                            if (CP_S.Name == "ВСЕ")
                            {
                            }
                            else
                            {

                                writer1.WriteStartElement("S_KAT");

                                //Сохраняем максимальный объем файлов в контейнере
                                writer1.WriteElementString("NAME", CP_S.Name);

                                writer1.WriteStartElement("SUB_KEY");

                                //Сохраняем максимальный объем файлов в контейнере
                                foreach (Key CP_K in CP_S.Keys)
                                {

                                    //Сохраняем максимальный объем файлов в контейнере
                                    writer1.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(CP_K.Value)));

                                }
                                //Закрываем корневой тег
                                writer1.WriteEndElement();


                                //Закрываем корневой тег
                                writer1.WriteEndElement();
                            }
                        }

                        //Закрываем корневой тег
                        writer1.WriteEndElement();



                        //Закрываем корневой тег
                        writer1.WriteEndElement();
                    }
                }

                writer1.WriteEndElement();



                //Закрываем корневой тег
                writer1.WriteEndElement();
                //Отпускаем поток записи
                writer1.Close();






                //=======================================================================================================
                //Загружаем товары из базы данных
                //=======================================================================================================

                FileInfo f = new FileInfo(_ProductDBPath + "\\ProductDB.csv");
                FileStream fileStream = new FileStream(f.FullName, FileMode.Create);
                StreamWriter sr = new StreamWriter(fileStream, Encoding.UTF8);

                string Line = "";

                try
                {

                    foreach (Product P in ProductListForPosting)
                    {
                        Line = P.CategoryOfProductName + "\t";
                        Line = Line + P.SubCategoryOfProductName + "\t";

                        Line = Line + P.datePost.ToString() + "\t";
                        Line = Line + P.HandBlock.ToString() + "\t";

                        Line = Line + P.IDURL + "\t";
                        Line = Line + P.Materials.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + "\t";

                        Line = Line + P.Prises.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + "\t";
                        Line = Line + P.Sizes.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + "\t";


                        foreach (int intten in P.prise)
                        {
                            Line = Line + intten + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.sellerText)
                        {
                            Line = Line + intten.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.sellerTextCleen)
                        {
                            Line = Line + intten.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.URLPhoto)
                        {
                            Line = Line + intten + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.FilePath)
                        {
                            Line = Line + intten + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";


                        sr.WriteLine(Line);
                    }


                }
                catch (Exception e1)
                {
                    MessageBox.Show("При сохранение товаров возникла ошибка: " + e1.ToString());
                }
                sr.Close();
                fileStream.Close();

                //=======================================================================================================

                //=======================================================================================================
                //Загружаем товары из базы данных
                //=======================================================================================================

                FileInfo f_UP = new FileInfo(_ProductDBPath + "\\ProductDBUnProcessed.csv");
                FileStream fileStream_UP = new FileStream(f_UP.FullName, FileMode.Create);
                StreamWriter sr_UP = new StreamWriter(fileStream_UP, Encoding.UTF8);

                Line = "";

                try
                {

                    foreach (Product P in ProductListSource)
                    {
                        Line = P.CategoryOfProductName + "\t";
                        Line = Line + P.SubCategoryOfProductName + "\t";

                        Line = Line + P.datePost.ToString() + "\t";
                        Line = Line + P.HandBlock.ToString() + "\t";

                        Line = Line + P.IDURL + "\t";
                        Line = Line + P.Materials.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + "\t";

                        Line = Line + P.Prises.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + "\t";
                        Line = Line + P.Sizes.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + "\t";


                        foreach (int intten in P.prise)
                        {
                            Line = Line + intten + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.sellerText)
                        {
                            Line = Line + intten.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.sellerTextCleen)
                        {
                            Line = Line + intten.Replace('\n', ' ').Replace('\t', ' ').Replace('\r', ' ') + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.URLPhoto)
                        {
                            Line = Line + intten + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";

                        foreach (string intten in P.FilePath)
                        {
                            Line = Line + intten + ",";
                        }
                        Line = Line.Substring(0, Line.Length - 1) + "\t";


                        sr_UP.WriteLine(Line);
                    }


                }
                catch (Exception e2)
                {
                    MessageBox.Show("При сохранение товаров возникла ошибка: " + e2.ToString());
                }
                sr_UP.Close();
                fileStream_UP.Close();

                //=======================================================================================================

            }
            catch (Exception e3)
            {
                MessageBox.Show("При сохранение товаров возникла ошибка: " + e3.ToString());
            }

        }


        // Экспорт данных из Excel-файла (не более 5 столбцов и любое количество строк <= 50.
        private int ExportExcel( string FilePath)
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
                Action S1 = () => label3.Text = "Загружено " + i + " из " + lastRow_1;
                label3.Invoke(S1);



                Product prod = new Product();

                prod.IDURL = ObjWorkSheet_1.Cells[i + 1, 4].Text.ToString();


                string lll = ObjWorkSheet_1.Cells[i + 1, 1].Text.ToString();
                // lll = "g:\\Job\\Education\\VKSMM\\ТЕСТ\\ФОТО\\" + lll.Substring(lll.IndexOf("\\") + 1);
                lll = _PhotoPath + "\\" + lll.Substring(lll.IndexOf("\\") + 1);//fbd.SelectedPath

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


                    if((www.Length>30)&&(www.IndexOf(".")>0))
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
                foreach (Product p in ProductListSourceBuffer)
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
                            ProductListSourceBuffer[iii].FilePath.Add(prod.FilePath[0]);
                            ProductListSourceBuffer[iii].URLPhoto.Add(prod.URLPhoto[0]);
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
                                ProductListSourceBuffer[iii].FilePath.Add(prod.FilePath[0]);
                                ProductListSourceBuffer[iii].URLPhoto.Add(prod.URLPhoto[0]);
                            }
                            else
                            {
                                imageNoExist.Add(prod.FilePath[0]);
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
                    ProductListSourceBuffer.Add(prod);
                    //dataGridView5.Rows.Add(prod.IDURL);
                    //listBox3.Items.Add(prod.IDURL);
                }
            }



            Action S2 = () => label3.Text = "Загружено " + lastRow_1 + " из " + lastRow_1+" "+(DateTime.Now-dateTime).ToString();
            label3.Invoke(S2);



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
        private int ExportProviderExcel()
        {
            try
            {

                //int countf = 888000;

                Excel.Application ObjWorkExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(_ProviderDir);//ofd.FileName_PhotoPath



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



                    try
                    {

                        CategoryOfProduct COFP = new CategoryOfProduct();
                        COFP.Name = ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString();

                        bool Reg = true;
                        int indexCat = -1;
                        int IOd = 0;
                        foreach (CategoryOfProduct C in mainCategoryList)
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





                                listBox1.Items.Add(COFP.Name);

                                mainCategoryList.Add(COFP);

                                string[] s = new string[2];

                                s[0] = COFP.Name;
                                s[1] = COFP.SubCategoty.Count.ToString();
                                dataGridView7.Rows.Add(s);
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

                                mainCategoryList[indexCat].Keys.Add(kmc);

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




        public void CreateExcel()
        {


                DirectoryInfo directory = new DirectoryInfo(ofd.FileName.Substring(0, ofd.FileName.LastIndexOf("\\")+1));






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
                foreach (Product p in ProductListForPosting)
                {

                    Action S2 = () => label15.Text = "Постов обработано " + t.ToString() + " из " + ProductListForPosting.Count.ToString();
                    label15.Invoke(S2);
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


                ProductListForPosting.Clear();
                treeView1.Nodes.Clear();
            listView2.Items.Clear();
            
        }

        // Импорт данных из Excel-файла (не более 5 столбцов и любое количество строк <= 50.
        //private int ImporExcel()
        //{
        //    // Выбрать путь и имя файла в диалоговом окне
        //    SaveFileDialog ofd = new SaveFileDialog();
        //    // Задаем расширение имени файла по умолчанию (открывается папка с программой)
        //    ofd.DefaultExt = "*.xls;*.xlsx";
        //    // Задаем строку фильтра имен файлов, которая определяет варианты
        //    ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
        //    // Задаем заголовок диалогового окна
        //    ofd.Title = "Выберите файл базы данных";
        //    if (!(ofd.ShowDialog() == DialogResult.OK)) // если файл БД не выбран -> Выход
        //        return 0;









        //    Excel.Application ObjWorkExcel = new Excel.Application();
        //    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);



        //    Excel.Worksheet ObjWorkSheet_1 = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1-й лист
        //    var lastCell_1 = ObjWorkSheet_1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку
        //                                                                                            // размеры базы
        //    int lastColumn_1 = (int)lastCell_1.Column;
        //    int lastRow_1 = (int)lastCell_1.Row;
        //    // Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 



        //    for (int i = 1; i < lastRow_1; i++) // по всем строкам
        //    {
        //        Product prod = new Product();

        //        prod.IDURL = ObjWorkSheet_1.Cells[i + 1, 4].Text.ToString();


        //        string lll = ObjWorkSheet_1.Cells[i + 1, 1].Text.ToString();
        //        lll = "g:\\Job\\Education\\VKSMM\\ТЕСТ\\ФОТО\\" + lll.Substring(lll.IndexOf("\\") + 1);
        //        prod.FilePath.Add(lll);


        //        prod.URLPhoto.Add(ObjWorkSheet_1.Cells[i + 1, 2].Text.ToString());
        //        prod.datePost = Convert.ToDateTime(ObjWorkSheet_1.Cells[i + 1, 3].Text.ToString());

        //        try
        //        {
        //            prod.prise[0] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 5].Text.ToString());
        //            prod.prise[1] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 6].Text.ToString());
        //            prod.prise[2] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 7].Text.ToString());
        //            prod.prise[3] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 8].Text.ToString());
        //            prod.prise[4] = Convert.ToInt32(ObjWorkSheet_1.Cells[i + 1, 9].Text.ToString());
        //        }
        //        catch { }


        //        try
        //        {
        //            string www = ObjWorkSheet_1.Cells[i + 1, 11].Text.ToString();
        //            while (www.IndexOf("\n") >= 0)
        //            {
        //                prod.sellerText.Add(www.Substring(0, www.IndexOf("\n")));
        //                www = www.Substring(www.IndexOf("\n") + 1);
        //            }
        //            prod.sellerText.Add(www);
        //        }
        //        catch { }


        //        prod.sellerTextCleen.Add(ObjWorkSheet_1.Cells[i + 1, 12].Text.ToString());

        //        bool get = false;
        //        int iii = 0;
        //        foreach (Product p in ProductListSource)
        //        {
        //            if (p.IDURL == prod.IDURL)
        //            {
        //                get = true;
        //                break;
        //            }
        //            iii++;
        //        }
        //        if (get)
        //        {
        //            ProductListSource[iii].FilePath.Add(prod.FilePath[0]);
        //            ProductListSource[iii].URLPhoto.Add(prod.URLPhoto[0]);
        //        }
        //        else
        //        {
        //            ProductListSource.Add(prod);
        //            //dataGridView5.Rows.Add(prod.IDURL);
        //            listBox3.Items.Add(prod.IDURL);
        //        }
        //    }








        //    Excel.Worksheet ObjWorkSheet_2 = (Excel.Worksheet)ObjWorkBook.Sheets[2]; //получить 1-й лист
        //    var lastCell_2 = ObjWorkSheet_2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//последнюю ячейку




        //    // размеры базы
        //    int lastColumn_2 = (int)lastCell_2.Column;
        //    int lastRow_2 = (int)lastCell_2.Row;
        //    // Перенос в промежуточный массив класса Form1: string[,] list = new string[50, 5]; 
        //    for (int i = 1; i < lastRow_2; i++) // по всем строкам
        //    {
        //        string[] LIN = new string[5];

        //        for (int j = 0; j < 5; j++) //по всем колонкам
        //        {
        //            LIN[j] = ObjWorkSheet_2.Cells[i + 1, j + 1].Text.ToString(); //считываем данные
        //        }

        //        dataGridView1.Rows.Add(LIN);
        //    }




        //    ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
        //    ObjWorkExcel.Quit(); // выйти из Excel
        //    GC.Collect(); // убрать за собой
        //    return 0;
        //}


        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



        List<string> imageNoExist = new List<string>();

        //Действия при выборе товара
        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (listBox1.SelectedIndex >= 0)
            //{
            //    ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = listBox1.SelectedItem.ToString();
            //}
            //if (listBox2.SelectedIndex >= 0)
            //{
            //    ProductListSource[listBox3.SelectedIndex].SubCategoryOfProductName = listBox2.SelectedItem.ToString();
            //}



            listBox1.SelectedItems.Clear();
            listBox2.SelectedItems.Clear();


            string Razmer = string.Empty;
            bool RazmerFind = false;


            //Если выбран товар то обрабатываем товар
            if (listBox3.SelectedItems.Count>0)
            {
                //Стираем данные от постовщика от старого товара
                dataGridView3.Rows.Clear();


                ProductListSource[listBox3.SelectedIndex].sellerTextCleen.Clear();

                //Буфер с новым описанием
                string stringpost = "";

                listBox7.Items.Clear();


                //Проходим по всем строчкам из описания
                for(int u = 0;u< ProductListSource[listBox3.SelectedIndex].sellerText.Count;u++)//listBox2.SelectedIndex
                {
                    string s = ProductListSource[listBox3.SelectedIndex].sellerText[u];

                    //В строчке должны быть данные
                    if (s.Length > 1)
                    {
                        //Добавляем строчку описания в грид
                        dataGridView3.Rows.Add(s);
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
                        if (u < ProductListSource[listBox3.SelectedIndex].sellerText.Count - 1)
                        {
                            if ((ProductListSource[listBox3.SelectedIndex].sellerText[u + 1].ToLower().IndexOf("рост") >= 0))
                            {
                                s = Razmer + " " + ProductListSource[listBox3.SelectedIndex].sellerText[u + 1];

                                u++;

                                //В строчке должны быть данные
                                if (ProductListSource[listBox3.SelectedIndex].sellerText[u].Length > 1)
                                {
                                    //Добавляем строчку описания в грид
                                    dataGridView3.Rows.Add(ProductListSource[listBox3.SelectedIndex].sellerText[u]);
                                }

                                RazmerFind = false;
                                Razmer = "";
                            }
                            else
                            {
                                if ((ProductListSource[listBox3.SelectedIndex].sellerText[u + 1].Length > 4))
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
                            foreach (ReplaceKeys k in Replace_Keys)
                            {
                                //Если ключ включен, то его исполняем
                                if (k.RegKey.IsActiv)
                                {
                                    //Регулярное выражение
                                    regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);

                                    if(regex.IsMatch(resultLine))
                                    {
                                        listBox7.Items.Add(k.RegKey.Value);
                                    }


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
                            foreach (ColorKeys k in Color_Keys)
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
                                        dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = k.color;
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
                                ProductListSource[listBox3.SelectedIndex].sellerTextCleen.Add(resultLine);
                                //Аккамулируем данные для поста
                                stringpost = stringpost + resultLine + "\r\n";
                            }
                            //Красим добавленную строчку 
                            dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor;
                        }
                    }
                }

                //Добавляем отредактированный пост на форму
                textBox3.Text = stringpost;




                //Чистим картинки в листбоксе
                imageList1.Images.Clear();
                listView1.Items.Clear();

                int ii = 0;
                foreach (string s in ProductListSource[listBox3.SelectedIndex].FilePath)
                {
                    if ((ProductListSource[listBox3.SelectedIndex].FilePath.Count == 1)
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

                                imageList1.Images.Add(new Bitmap(s));
                                listView1.Items.Add(new ListViewItem(s, ii));
                                ii++;
                            }
                            else
                            {
                                imageNoExist.Add(s);
                            }
                        }
                        catch (Exception ee)
                        {
                            MessageBox.Show(ee.ToString());
                        }
                    }
                }



                


                //========================== Блок с автоподбором категорий ====================================================
                int i = 0;
                int j = 0;
                //Регулярное выражение
                Regex regexCat;
                //Перебираем категории товара
                foreach (CategoryOfProduct c in mainCategoryList)
                {
                    //Флаг обнаружения категории
                    bool regincat = false;

                    if (ProductListSource[listBox3.SelectedIndex].CategoryOfProductName == c.Name)
                    {
                        listBox1.SelectedIndex = i;
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
                        if (ProductListSource[listBox3.SelectedIndex].CategoryOfProductName == "ВСЕ")
                        {
                            ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = c.Name;
                        }
                        if (!ProductListSource[listBox3.SelectedIndex].HandBlock)
                        {
                            //Выделяем категорию
                            listBox1.SelectedIndex = i;
                        }
                        j = 0;
                        //Перебираем подкатегории выбранной категории
                        foreach (SubCategoryOfProduct s in c.SubCategoty)
                        {
                            //
                            bool reginsubcat = false;

                            if (ProductListSource[listBox3.SelectedIndex].SubCategoryOfProductName == s.Name)
                            {
                                listBox2.SelectedIndex = j;
                            }

                                foreach (Key k in s.Keys)
                            {
                                if (k.IsActiv)
                                {
                                    regexCat = new Regex(k.Value, RegexOptions.IgnoreCase);

                                    reginsubcat = regexCat.IsMatch(stringpost) || reginsubcat;
                                }
                            }

                            if(reginsubcat)
                            {
                                if (ProductListSource[listBox3.SelectedIndex].SubCategoryOfProductName == "ВСЕ")
                                {
                                    ProductListSource[listBox3.SelectedIndex].SubCategoryOfProductName = s.Name;
                                }

                                if (!ProductListSource[listBox3.SelectedIndex].HandBlock)
                                {
                                    listBox2.SelectedIndex = j;
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






                numericUpDown2.Value = ProductListSource[listBox3.SelectedIndex].prise[0];
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
            r.Action = ActionDecoder(comboBox3.Text);
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



        /// <summary>
        /// Путь к месту запуска программы
        /// </summary>
        public string _path = "";





        public int selectedIndexCategory = -1;
        public int selectedIndexSubCategory = -1;
        public int selectedIndexReplace = -1;
        public int selectedIndexColor = -1;

     

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



        Regex regex = new Regex(@"туп(\w*)", RegexOptions.IgnoreCase);
       // Regex.IsMatch(email, pattern, RegexOptions.IgnoreCase)

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
           // if (dataGridView5.SelectedRows.Count > 0)
            {
                dataGridView3.Rows.Clear();
                //listBox1.Items.Add(ProductListSource[listBox2.SelectedIndex].sellerText[0]);
                // string[] ssss = new string[ProductListSource[listBox2.SelectedIndex].sellerText.Count]; 
                string stringpost = "";

                int i = 0;
                foreach(Product p in ProductListSource)
                {
                  //  if(p.IDURL== dataGridView5.SelectedRows[0].Cells[0].Value.ToString())
                    {
                        break;
                    }
                    i++;
                }



                foreach (string s in ProductListSource[i].sellerText)//listBox2.SelectedIndex
                {
                    dataGridView3.Rows.Add(s);

                    
                    dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = Color.Red;
                    

                    foreach (DataGridViewRow r in dataGridView2.Rows)
                    {
                        if (s.IndexOf(r.Cells[0].Value.ToString()) >= 0)
                        {
                            dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor = r.DefaultCellStyle.BackColor;
                            break;
                        }
                    }
                    foreach (DataGridViewRow r in dataGridView2.Rows)
                    {
                        if (dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor == Color.LightGreen)
                        {
                            stringpost = stringpost + s + "\r\n";
                            break;
                        }
                    }


                    dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.SelectionBackColor = dataGridView3.Rows[dataGridView3.Rows.Count - 1].DefaultCellStyle.BackColor;
                }

                textBox7.Text = stringpost;

                imageList1.Images.Clear();
                listView1.Items.Clear();
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
            r.Action = ActionDecoder(comboBox1.Text);
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

        private void button1_Click_1(object sender, EventArgs e)
        {

            imageNoExist.Clear();


            // Выбрать путь и имя файла в диалоговом окне
            OpenFileDialog ofd = new OpenFileDialog();
            // Задаем расширение имени файла по умолчанию (открывается папка с программой)
            ofd.DefaultExt = "*.xls;*.xlsx";
            // Задаем строку фильтра имен файлов, которая определяет варианты
            ofd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            // Задаем заголовок диалогового окна
            ofd.Title = "Выберите файл базы данных";
            if ((ofd.ShowDialog() == DialogResult.OK)) // если файл БД не выбран -> Выход
            {


                //// Выбрать путь и имя файла в диалоговом окне
                //FolderBrowserDialog fbd = new FolderBrowserDialog();

                //if (!(fbd.ShowDialog() == DialogResult.OK)) // если файл БД не выбран -> Выход
                //    return 0;



                ExportExcel(ofd.FileName);

                MessageBox.Show("Данные загружены!");


                //t1 = new Thread(_PlayerNAUDIO.PlayBack_Codes);
                //t1.Priority = ThreadPriority.Lowest;
                //t1.Start();
                //t2 = new Thread(_PlayerNAUDIO.PlayBack_Codes);
                //t2.Priority = ThreadPriority.Lowest;
                //t2.Start();
                //t3 = new Thread(_PlayerNAUDIO.PlayBack_Codes);
                //t3.Priority = ThreadPriority.Lowest;
                //t3.Start();
                //t4 = new Thread(_PlayerNAUDIO.PlayBack_Codes);
                //t4.Priority = ThreadPriority.Lowest;
                //t4.Start();



                foreach (Product P in ProductListSource)
                {
                    listBox3.Items.Add(P.IDURL);
                }


                string sL = "При загрузке отсутствуют следующие изображения: \r\n";
                foreach (string L in imageNoExist)
                {
                    sL = sL + L + "\r\n";
                }

                MessageBox.Show(sL);
            }

            UpdatePostavshikov();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();

            if ((listBox3.SelectedIndex >= 0)&&(listBox1.SelectedIndex>=0))
            {
                ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = listBox1.SelectedItem.ToString();
            }
            
            try
            {
                foreach (SubCategoryOfProduct sub in mainCategoryList[listBox1.SelectedIndex].SubCategoty)
                {
                    listBox2.Items.Add(sub.Name);
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
                    comboBox3.Text = ActionCoder(Replace_Keys[selectedIndexReplace].Action);
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
                        comboBox1.Text = ActionCoder(Color_Keys[selectedIndexColor].Action);


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
                Replace_Keys[i].Action = ActionDecoder(comboBox3.Text);

                AddGroup(Replace_Keys[i]);
            }

        }

        private void button14_Click(object sender, EventArgs e)
        {
            Color_Keys[selectedIndexColor].RegKey.Value = textBox4.Text;
            Color_Keys[selectedIndexColor].Action = ActionDecoder(comboBox1.Text);

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


            RemovedIndexes = ProcessingProducts(listBox3.SelectedIndex);


            for (int iq = RemovedIndexes.Count - 1; iq >= 0; iq--)
            {
                listBox3.Items.RemoveAt(RemovedIndexes[iq]);
                ProductListSource.RemoveAt(RemovedIndexes[iq]);
            }

            UpdatePostavshikov();

            listBox3.SelectedItems.Clear();

            textBox3.Text = "";
            listView1.Items.Clear();
            numericUpDown2.Value = 0;
            dataGridView3.Rows.Clear();

            //try
            //{
            //    treeView1.SelectedNode = treeView1.Nodes[0];
            //}
            //catch
            //{ }


            //bool goodsProcessed = true;

            //foreach(DataGridViewRow r in dataGridView1.Rows)
            //{
            //    if((r.DefaultCellStyle.BackColor==Color.LightBlue)|| (r.DefaultCellStyle.BackColor == Color.LightGreen))
            //    {

            //    }
            //    else
            //    {
            //        goodsProcessed = false;
            //    }
            //}



            //if (goodsProcessed)
            //{
            //    try
            //    {
            //        int index = listBox3.SelectedIndex;

            //        ProductListForPosting.Add(ProductListSource[index]);

            //        AddToTreeView(ProductListSource[index], ProductListForPosting.Count - 1);

            //        ProductListSource.RemoveAt(index);

            //        try
            //        {
            //            listBox3.Items.Remove(listBox3.SelectedItem);
            //        }
            //        catch
            //        { }

            //        textBox3.Text = "";
            //        listView1.Items.Clear();
            //        numericUpDown2.Value = 0;

            //    }
            //    catch (Exception we)
            //    {
            //        MessageBox.Show("Выберите товар!");
            //    }
            //}
            //else
            //{
            //    MessageBox.Show("Не все строчки описания товара обработаны!");
            //}
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




            listBox3.SelectedItems.Clear();

            textBox3.Text = "";
            listView1.Items.Clear();
            numericUpDown2.Value = 0;
            dataGridView3.Rows.Clear();


            List<int> RemovedIndexes = new List<int>();


            RemovedIndexes = ProcessingProductsAll();


            for (int iq = RemovedIndexes.Count-1; iq>=0;iq-- )
            {
                listBox3.Items.RemoveAt(RemovedIndexes[iq]);
                ProductListSource.RemoveAt(RemovedIndexes[iq]);
            }

            UpdatePostavshikov();
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


        public List<int> ProcessingProductsAll()
        {
            int index = 0;
            //int indexTV = 0;
            int indexTV = ProductListForPosting.Count;

            List<int> RemovedIndexes = new List<int>();

            foreach (Product P in ProductListSource)
            {
                if (((P.IDURL.IndexOf(comboBox5.Text) >= 0) || (comboBox5.Text == "ВСЕ"))
                    &&((P.CategoryOfProductName== comboBox4.Text) ||(comboBox4.Text == "ВСЕ")))
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

                        if ((s.ToLower().IndexOf("размер") >= 0)||(s.ToLower().IndexOf("разм.") >= 0) || (s.ToLower().IndexOf("рост") >= 0) || (s.ToLower().IndexOf("opct") >= 0))
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
                                foreach (ReplaceKeys k in Replace_Keys)
                                {
                                    //Если ключ включен, то его исполняем
                                    if (k.RegKey.IsActiv)
                                    {
                                        //Регулярное выражение
                                        regex = new Regex(k.RegKey.Value, RegexOptions.IgnoreCase);

                                        if(regex.IsMatch(resultLine))
                                        {
                                            P.logRegularExpression.Add(k.RegKey.Value);
                                        }


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
                                foreach (ColorKeys k in Color_Keys)
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

                    //if (i == 0)
                    //{ //MessageBox.Show("Внимание такой категории не существует!");
                    //  //break;
                    //}
                    ////============================================================= временное решение ===============================================================================
                    //========================== Блок с автоподбором категорий ====================================================
                    int i = 0;
                    int j = 0;
                    //Регулярное выражение
                    Regex regexCat;
                    //Перебираем категории товара
                    foreach (CategoryOfProduct c in mainCategoryList)
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

                        ProductListForPosting.Add(P);

                        AddToTreeView(P, indexTV);

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

        public List<int> ProcessingProducts(int indx)
        {
            int index = 0;
            //int indexTV = 0;
            int indexTV = ProductListForPosting.Count;

            List<int> RemovedIndexes = new List<int>();

            Product P = ProductListSource[indx];
            {
                if ((P.IDURL.IndexOf(comboBox5.Text) >= 0) || (comboBox5.Text == "ВСЕ"))
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
                                foreach (ReplaceKeys k in Replace_Keys)
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
                                foreach (ColorKeys k in Color_Keys)
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

                    //if (i == 0)
                    //{ //MessageBox.Show("Внимание такой категории не существует!");
                    //  //break;
                    //}
                    ////============================================================= временное решение ===============================================================================
                    //========================== Блок с автоподбором категорий ====================================================
                    int i = 0;
                    int j = 0;
                    //Регулярное выражение
                    Regex regexCat;
                    //Перебираем категории товара
                    foreach (CategoryOfProduct c in mainCategoryList)
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



                    //if (isCat)
                    //{
                    //    isSub = true;
                    //    if (P.SubCategoryOfProductName == "ВСЕ")
                    //    {
                    //        P.SubCategoryOfProductName = "ВСЕ";
                    //    }
                    //}

                    //===============================================================================================================================================================



                   // if (isCat && isSub && isStop)//&& goodsProcessed
                    {

                        ProductListForPosting.Add(P);

                        AddToTreeView(P, indexTV);

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
                imageList1.Images.Clear();
                listView2.Items.Clear();

                int ii = 0;
                foreach (string s in ProductListForPosting[indexOfProduct].FilePath)
                {
                    try
                    {
                        if (s.Length > 3)
                        {
                            imageList1.Images.Add(new Bitmap(s));
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
                imageList1.Images.Clear();
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
                                    imageList1.Images.Add(new Bitmap(s));
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
                imageList1.Images.Clear();
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
                                    imageList1.Images.Add(new Bitmap(ProductListForPosting[indexOfProduct].FilePath[0]));
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

        public List<string> swownProductInListView = new List<string>();

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

        // Выбрать путь и имя файла в диалоговом окне
        SaveFileDialog ofd = new SaveFileDialog();


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
            CreateExcel();
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
            if ((listBox3.SelectedIndex >= 0) && (listBox2.SelectedIndex >= 0))
            {
                // ProductListSource[listBox3.SelectedIndex].CategoryOfProductName = listBox1.SelectedItem.ToString();

                ProductListSource[listBox3.SelectedIndex].SubCategoryOfProductName = listBox2.SelectedItem.ToString();

                ProductListSource[listBox3.SelectedIndex].HandBlock = true;
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
                        Filtr[2] = ActionCoder(RK.Action);

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
                    Filtr[2] = ActionCoder(RK.Action);

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

        private void button15_Click(object sender, EventArgs e)
        {
            Thread_Dir_Processing = new Thread(Thread_Dir_Processing_Code);
            Thread_Dir_Processing.Start();
        }

        public void Thread_Dir_Processing_Code()
        {
            Action S1 = () => button1.Enabled = false;
            button1.Invoke(S1);

            DirectoryInfo d = new DirectoryInfo(_InputPath);

            if (!d.Exists)
            {
                MessageBox.Show("Директория " + _InputPath + " не существует!");
            }
            else
            {

                int cf = d.GetFiles().Length;
                int pf = 0;

                imageNoExist.Clear();


                foreach (FileInfo f in d.GetFiles())
                {
                    try
                    {
                        Action S2 = () => label11.Text = "Файлов обработано " + pf.ToString() + " из " + cf.ToString();
                        label11.Invoke(S2);




                        ExportExcel(f.FullName);



                        f.Delete();
                        pf++;
                    }
                    catch
                    { }
                }





                foreach (Product p in ProductListSourceBuffer)
                {
                    ProductListSource.Add(p);
                }


                ProductListSourceBuffer.Clear();





                string sL = "При загрузке отсутствуют следующие изображения: \r\n";
                foreach (string L in imageNoExist)
                {
                    sL = sL + L + "\r\n";
                }

                MessageBox.Show(sL);






                listBox3.Items.Clear();

                int i = 1;



                foreach (Product P in ProductListSource)
                {
                    listBox3.Items.Add(i);
                    i++;
                }

                UpdatePostavshikov();

                Action S3 = () => button1.Enabled = true;
                button1.Invoke(S3);
            }

        }


        private void UpdatePostavshikov()
        {
            comboBox5.Items.Clear();

            comboBox5.Items.Add("ВСЕ");
            comboBox5.Text = "ВСЕ";

            foreach (Product P in ProductListSource)
            {
                bool gift = true;

                string L = P.IDURL.Substring(P.IDURL.IndexOf("/id") + 1);

                L = L.Substring(0, L.IndexOf("?"));


                foreach (object c in comboBox5.Items)
                {
                    if (c.ToString() == L)
                    {
                        gift = false;
                    }
                }

                if (gift)
                {
                    comboBox5.Items.Add(L);
                }
            }
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
                        Filtr[2] = ActionCoder(RK.Action);

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
                    Filtr[2] = ActionCoder(RK.Action);

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
                    listBox1.Items.Add(newCategory.Name);
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
            listBox1.Items.RemoveAt(u);
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
                listBox1.Items[u] = textBox6.Text;

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

                    if (listBox1.SelectedIndex == u)
                    {
                        listBox2.Items.Add(newSubCP.Name);
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

            if (listBox1.SelectedIndex == u)
            {
                listBox2.Items.RemoveAt(w);
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

                if (listBox1.SelectedIndex == u)
                {
                    listBox2.Items[w] = textBox9.Text;
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
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            foreach (CategoryOfProduct c in mainCategoryList)
            {
                listBox1.Items.Add(c.Name);
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

        }
    }
}

