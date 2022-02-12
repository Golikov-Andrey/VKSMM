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
using VKSMM.ThredsCode;//Файл с классами потока

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
            Thread_Load_Data = new Thread(Core.Thread_Provider_Excel_Code);
            Thread_Load_Data.Start(this);

        }

        /// <summary>
        /// Процесс загрузки данных при запуске программы
        /// </summary>
        public Thread Thread_Load_Data;

        /// <summary>
        /// Ссылка на основное окно
        /// </summary>
        public MainForm mainForm;

    }
}
