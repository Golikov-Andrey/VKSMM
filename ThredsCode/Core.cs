using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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

                Stuff.UpdatePostavshikov(mainForm);

                //Action S3 = () => button1.Enabled = true;
                //button1.Invoke(S3);
            }

        }

    }
}
