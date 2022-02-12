using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using VKSMM.ModelClasses;//Файл с классами моделей данных

namespace VKSMM.StuffClasses
{
    /// <summary>
    /// Класс записи данных программы на жесткий диск
    /// </summary>
    class ConfigWriter
    {
        public static void writeDirToConfigFile(MainForm mainForm, XmlWriter writerConfigXML)
        {
            //Сохраняем максимальный объем файлов в контейнере
            writerConfigXML.WriteElementString("PRODUCT_DB_DIR", mainForm._ProductDBPath);

            //Сохраняем максимальный объем файлов в контейнере
            writerConfigXML.WriteElementString("PHOTO_DIR", mainForm._PhotoPath);

            //Сохраняем максимальный объем файлов в контейнере
            writerConfigXML.WriteElementString("INPUT_DIR", mainForm._InputPath);

            //Сохраняем максимальный объем файлов в контейнере
            writerConfigXML.WriteElementString("PROVIDER_DIR", mainForm._ProviderDir);

        }

        public static void writeReplaceKeyToConfigFile(MainForm mainForm, XmlWriter writerConfigXML)
        {
            //Корневой тег
            writerConfigXML.WriteStartElement("REPLACE_KEYS");

            foreach (ReplaceKeys RK in mainForm.Replace_Keys)
            {
                //Корневой тег
                writerConfigXML.WriteStartElement("R_KEY");

                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(RK.RegKey.Value)));
                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("VALUE", RK.NewValue);
                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("MODE", Stuff.ActionCoder(RK.Action));
                //Сохраняем максимальный объем файлов в контейнере
                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("GROUP", RK.GroupValue);

                //Закрываем корневой тег
                writerConfigXML.WriteEndElement();
            }
            writerConfigXML.WriteEndElement();
        }

        public static void writeColorKeyToConfigFile(MainForm mainForm, XmlWriter writerConfigXML)
        {
            writerConfigXML.WriteStartElement("COLOR_KEYS");

            foreach (ColorKeys CK in mainForm.Color_Keys)
            {
                //Корневой тег
                writerConfigXML.WriteStartElement("C_KEY");

                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(CK.RegKey.Value)));
                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("MODE", Stuff.ActionCoder(CK.Action));
                //Сохраняем максимальный объем файлов в контейнере
                writerConfigXML.WriteElementString("COLOR", CK.color.Name);
                //Сохраняем максимальный объем файлов в контейнере

                //Закрываем корневой тег
                writerConfigXML.WriteEndElement();
            }

            writerConfigXML.WriteEndElement();

        }

        public static void writeCategoryProductToConfigFile(MainForm mainForm, XmlWriter writerConfigXML)
        {
            writerConfigXML.WriteStartElement("CATEGORY_PRODUCT");

            foreach (CategoryOfProduct CP in mainForm.mainCategoryList)
            {
                if (CP.isProvider)
                {

                    //Корневой тег
                    writerConfigXML.WriteStartElement("CATEG_KEYS");

                    //Сохраняем максимальный объем файлов в контейнере
                    writerConfigXML.WriteElementString("NAME", CP.Name);

                    writerConfigXML.WriteStartElement("KEYS");

                    foreach (Key CP_K in CP.Keys)
                    {
                        if (CP_K.isProvider)
                        {

                            writerConfigXML.WriteStartElement("CAT_KEY");

                            //Сохраняем максимальный объем файлов в контейнере
                            writerConfigXML.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(CP_K.Value)));

                            //Закрываем корневой тег
                            writerConfigXML.WriteEndElement();
                        }
                    }

                    //Закрываем корневой тег
                    writerConfigXML.WriteEndElement();


                    writerConfigXML.WriteStartElement("SUB_KATS");

                    foreach (SubCategoryOfProduct CP_S in CP.SubCategoty)
                    {
                        if (CP_S.Name == "ВСЕ")
                        {
                        }
                        else
                        {

                            writerConfigXML.WriteStartElement("S_KAT");

                            //Сохраняем максимальный объем файлов в контейнере
                            writerConfigXML.WriteElementString("NAME", CP_S.Name);

                            writerConfigXML.WriteStartElement("SUB_KEY");

                            //Сохраняем максимальный объем файлов в контейнере
                            foreach (Key CP_K in CP_S.Keys)
                            {

                                //Сохраняем максимальный объем файлов в контейнере
                                writerConfigXML.WriteElementString("KEY", Convert.ToBase64String(Encoding.Unicode.GetBytes(CP_K.Value)));

                            }
                            //Закрываем корневой тег
                            writerConfigXML.WriteEndElement();


                            //Закрываем корневой тег
                            writerConfigXML.WriteEndElement();
                        }
                    }

                    //Закрываем корневой тег
                    writerConfigXML.WriteEndElement();



                    //Закрываем корневой тег
                    writerConfigXML.WriteEndElement();
                }
            }

            writerConfigXML.WriteEndElement();

        }

        public static void writeProductDB(MainForm mainForm)
        {
            //=======================================================================================================
            FileInfo f = new FileInfo(mainForm._ProductDBPath + "\\ProductDB.csv");
            FileStream fileStream = new FileStream(f.FullName, FileMode.Create);
            StreamWriter sr = new StreamWriter(fileStream, Encoding.UTF8);

            string Line = "";

            try
            {

                foreach (Product P in mainForm.ProductListForPosting)
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

                    foreach (string intten in P.logRegularExpression)
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

        }

        public static void writeProductDBUnProcessed(MainForm mainForm)
        {

            //=======================================================================================================
            FileInfo f_UP = new FileInfo(mainForm._ProductDBPath + "\\ProductDBUnProcessed.csv");
            FileStream fileStream_UP = new FileStream(f_UP.FullName, FileMode.Create);
            StreamWriter sr_UP = new StreamWriter(fileStream_UP, Encoding.UTF8);

            string Line = "";

            try
            {

                foreach (Product P in mainForm.productListSource)
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

        }

    }
}
