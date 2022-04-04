using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VKSMM
{
    /// <summary>
    /// Окно ввода пароля Администратора
    /// </summary>
    public partial class EnterPassForm : Form
    {
        public EnterPassForm()
        {
            InitializeComponent();
        }


        /// <summary>
        /// Административный пароль
        /// </summary>
        public string administrativPassWord = "CD5QH97J67";

        /// <summary>
        /// Ссылка на основное окно
        /// </summary>
        public MainForm mainForm;

        /// <summary>
        /// Проверка пароля
        /// </summary>
        private void buttonTestPass_Click(object sender, EventArgs e)
        {
            //Проверяем пароль
            if (textBox1.Text == administrativPassWord)//true)
            {
                //Поднимаем флаг административного режима
                mainForm.administrativPassEnter = true;

                //Включаем отключенные элементы управления 
                mainForm.groupBoxLoadProduct.Visible = true;
                mainForm.groupBoxKeyManager.Visible = true;
                mainForm.groupBoxMainCategory.Visible = true;
                mainForm.groupBoxSubCategory.Visible = true;
                mainForm.providerDataGrid.Visible = true;

                //Закрываем окно ввода пароля
                this.Close();
            }
            else
            {
                MessageBox.Show("Пароль введен неверно!");
            }
        }

        /// <summary>
        /// Загрузка формы без административной части
        /// </summary>
        private void buttonIgnorePass_Click(object sender, EventArgs e)
        {
            //Закрываем окно ввода пароля
            this.Close();
        }
    }
}
