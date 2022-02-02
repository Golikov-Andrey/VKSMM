using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using VKSMM.ModelClasses;//Файл с классами моделей данных

namespace VKSMM.StuffClasses
{
    class ConfigReader
    {

        public static void readDirFromConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {
            try
            {
                nodeList = root.SelectNodes("PRODUCT_DB_DIR");//Считываем настройки ключей замены

                mainForm._ProductDBPath = nodeList[0].InnerText;
            }
            catch
            {

            }

            try
            {
                nodeList = root.SelectNodes("PHOTO_DIR");//Считываем настройки ключей замены

                mainForm._PhotoPath = nodeList[0].InnerText;
            }
            catch
            {

            }
            try
            {
                nodeList = root.SelectNodes("INPUT_DIR");//Считываем настройки ключей замены

                mainForm._InputPath = nodeList[0].InnerText;
            }
            catch
            {

            }

            try
            {
                nodeList = root.SelectNodes("PROVIDER_DIR");//Считываем настройки ключей замены

                mainForm._ProviderDir = nodeList[0].InnerText;
            }
            catch
            {

            }

        }

        public static void readReplaceKeyFromConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {
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

                    Filtr[3] = mainForm.GUIDReplaceKey.ToString();

                    Filtr[4] = PL.ChildNodes[3].InnerText;

                    mainForm.GUIDReplaceKey++;

                    //Добавляем правило в ГРИД
                    mainForm.dataGridView2.Rows.Add(Filtr);

                    //================================================================

                    //=========== Конструктор для добавления в датагрид =============
                    //Создаем экземпляр ключа
                    Key k = new Key();
                    //Значение ключа
                    k.Value = Filtr[0].Replace("\r", "").Replace("\n", "");
                    //Флаг активности ключа
                    k.IsActiv = true;

                    //Класс замены
                    ReplaceKeys r = new ReplaceKeys();
                    //Действие связанное с заменой 3-ка просто замена
                    r.Action = Stuff.ActionDecoder(Filtr[2]);
                    //Ключ приыязанный к классу замены
                    r.RegKey = k;
                    //Значение замены
                    r.NewValue = Filtr[1].Replace("\r", "").Replace("\n", "");

                    try
                    {
                        r.GroupValue = PL.ChildNodes[3].InnerText;
                    }
                    catch { }

                    mainForm.AddGroup(r);

                    //Добавляем ключ в пул замен
                    mainForm.Replace_Keys.Add(r);
                    //================================================================
                }
            }
            catch { }

        }

        public static void readColorKeyFromConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {
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
                    mainForm.dataGridView4.Rows.Add(Filtr);
                    //Окрашываем строчку
                    mainForm.dataGridView4.Rows[mainForm.dataGridView4.Rows.Count - 1].DefaultCellStyle.BackColor = Color.FromName(PL.ChildNodes[2].InnerText);
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
                    r.Action = Stuff.ActionDecoder(Filtr[1]);
                    //Ключ приыязанный к классу замены
                    r.RegKey = k;
                    //Значение замены
                    r.color = Color.FromName(PL.ChildNodes[2].InnerText);

                    //Добавляем ключ в пул замен
                    mainForm.Color_Keys.Add(r);
                    //================================================================
                }
            }
            catch { }

        }

        public static void readCategoryProductFromConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {
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


                    mainForm.listBox1.Items.Add(COFP.Name);

                    mainForm.mainCategoryList.Add(COFP);

                    string[] s = new string[2];

                    s[0] = COFP.Name;
                    s[1] = COFP.SubCategoty.Count.ToString();
                    mainForm.dataGridView7.Rows.Add(s);
                }


            }
            catch { }


        }

        public static void readProductDB(MainForm mainForm)
        {
            //=======================================================================================================
            //Загружаем товары из базы данных
            //=======================================================================================================

            FileInfo f = new FileInfo(mainForm._ProductDBPath + "\\ProductDB.csv");
            FileStream fileStream = new FileStream(f.FullName, FileMode.Open);
            StreamReader sr = new StreamReader(fileStream, Encoding.UTF8);

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


                    PN.prise = Stuff.ConvertMassToInt(MassLine[8].Split(new char[] { ',' }));

                    PN.sellerText = Stuff.ConvertMassToList(MassLine[9].Split(new char[] { ',' }));

                    PN.sellerTextCleen = Stuff.ConvertMassToList(MassLine[10].Split(new char[] { ',' }));

                    PN.URLPhoto = Stuff.ConvertMassToList(MassLine[11].Split(new char[] { ',' }));

                    PN.FilePath = Stuff.ConvertMassToList(MassLine[12].Split(new char[] { ',' }));

                    mainForm.ProductListForPosting.Add(PN);

                    mainForm.AddToTreeView(PN, _it);
                    _it++;
                }
                catch
                {

                }
            }

            sr.Close();
            fileStream.Close();

            //=======================================================================================================

        }

        public static void readProductDBUnProcessed(MainForm mainForm)
        {
            //=======================================================================================================
            //Загружаем товары из базы данных
            //=======================================================================================================

            FileInfo f_UP = new FileInfo(mainForm._ProductDBPath + "\\ProductDBUnProcessed.csv");
            FileStream fileStream_UP = new FileStream(f_UP.FullName, FileMode.Open);
            StreamReader sr_UP = new StreamReader(fileStream_UP, Encoding.UTF8);


            string Line = "";


            int _it = 0;

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


                    PN.prise = Stuff.ConvertMassToInt(MassLine[8].Split(new char[] { ',' }));

                    PN.sellerText = Stuff.ConvertMassToList(MassLine[9].Split(new char[] { ',' }));

                    PN.sellerTextCleen = Stuff.ConvertMassToList(MassLine[10].Split(new char[] { ',' }));

                    PN.URLPhoto = Stuff.ConvertMassToList(MassLine[11].Split(new char[] { ',' }));

                    PN.FilePath = Stuff.ConvertMassToList(MassLine[12].Split(new char[] { ',' }));

                    mainForm.ProductListSource.Add(PN);

                    mainForm.listBox3.Items.Add(PN.IDURL);
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

    }
}
