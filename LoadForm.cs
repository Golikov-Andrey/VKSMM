using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using VKSMM.StuffClasses;//Файл с классами вспомогательных методов

namespace VKSMM
{
    /// <summary>
    /// Окно отображения загрузки данных при запуске программы
    /// </summary>
    public partial class LoadForm : Form
    {
        public LoadForm()
        {
            InitializeComponent();

            //Запускаем поток конвертации файла с постовщиками во внутренний формат
            Thread_Load_Data = new Thread(Thread_Load_Data_Code);
            Thread_Load_Data.Start();

        }

        /// <summary>
        /// Процесс загрузки данных при запуске программы
        /// </summary>
        public Thread Thread_Load_Data;

        /// <summary>
        /// Ссылка на основное окно
        /// </summary>
        public MainForm mainForm;

        public void Thread_Load_Data_Code()
        {
            Stuff.ExportProviderExcel(mainForm, this);

            Action S1 = () => this.Close();
            this.Invoke(S1);            
        }
    }
}
