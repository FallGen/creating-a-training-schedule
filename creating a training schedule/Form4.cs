using System;
using System.Windows.Forms;

namespace creating_a_training_schedule
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
            dataGridView1.RowCount = 14;
        }

        bool close = false;

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (textBox1.Text == "" || textBox1.Text == null)
                    throw new Exception("не задано имя присета времени по умолчанию");
            }
            catch (Exception E)
            {
                if (close) e.Cancel = false;
                else
                {
                    e.Cancel = true;
                    textBox1.Focus();
                    MessageBox.Show(E.Message, "предупреждение!");
                }
            }
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            close = true;
        }
    }
}
