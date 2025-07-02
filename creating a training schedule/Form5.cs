using System;
using System.Windows.Forms;

namespace creating_a_training_schedule
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        bool close = false;
        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (comboBox1.SelectedIndex == -1 || comboBox1.Items == null)
                    throw new Exception("не выбран имя присета времени по умолчанию");
            }
            catch (Exception E)
            {
                if (close)
                    e.Cancel = false;
                else
                {
                    e.Cancel = true;
                    comboBox1.Focus();
                    MessageBox.Show(E.Message, "предупреждение!");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            close = true;
        }
    }
}
