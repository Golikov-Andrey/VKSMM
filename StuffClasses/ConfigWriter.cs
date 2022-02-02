using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using VKSMM.ModelClasses;//Файл с классами моделей данных

namespace VKSMM.StuffClasses
{
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

        public static void writeReplaceKeyToConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {

        }

        public static void writeColorKeyToConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {

        }

        public static void writeCategoryProductToConfigFile(MainForm mainForm, XmlNodeList nodeList, XmlNode root)
        {

        }

        public static void writeProductDB(MainForm mainForm)
        {
            //=======================================================================================================

        }

        public static void writeProductDBUnProcessed(MainForm mainForm)
        {

            //=======================================================================================================
        }

    }
}
