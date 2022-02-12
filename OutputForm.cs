using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VKSMM.ThredsCode;//

namespace VKSMM
{
    public partial class OutputForm : Form
    {
        public OutputForm()
        {
            InitializeComponent();

            Thread_Create_XLS_Processing = new Thread(Core.Thread_Create_Excel_Code);
            Thread_Create_XLS_Processing.Start(this);

        }

        /// <summary>
        /// Процесс выгрузки данных 
        /// </summary>
        public Thread Thread_Create_XLS_Processing;

        /// <summary>
        /// Ссылка на основное окно
        /// </summary>
        public MainForm mainForm;

        /// <summary>
        /// Путь к выгружаемому посту, XML файл 
        /// </summary>
        public string outputFilePath = "";

        //public void Thread_Create_XLS_Processing_Code()
        //{
        //    Stuff.CreateExcel(this, outputFilePath);

        //    Action S1 = () => this.Close();
        //    this.Invoke(S1);
        //}
    }
}
