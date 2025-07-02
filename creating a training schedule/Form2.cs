using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace creating_a_training_schedule
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        bool close = false;

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (textBox1.Text == "" || textBox1.Text == null)
                    throw new Exception("не заполнены данные");
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

        private void button2_Click(object sender, EventArgs e)
        {
            close = true;
        }
    }
}
