using System;
using System.Windows.Forms;

namespace creating_a_training_schedule
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        bool close = false;
        private void chenge_rb(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                comboBox1.Visible = false;
                label2.Visible = false;
            }
            else
            {
                comboBox1.Visible = true;
                label2.Visible = true;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
                comboBox2.Visible = true;
            else comboBox2.Visible = false;
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                try
                {
                    throw new Exception("не выбран параметр создание нового расписания");
                }
                catch (Exception E)
                {
                    if (close)
                        e.Cancel = false;
                    else
                    {
                        e.Cancel = true;
                        groupBox1.Focus();
                        MessageBox.Show(E.Message, "предупреждение!");
                    }
                }
            }
            //else
            if (comboBox2.Visible)
            {
                try
                {
                    if (comboBox2.SelectedIndex == -1 || comboBox2.Items == null)
                        throw new Exception("не выбрано время тренировок по умолчанию");
                }
                catch (Exception E)
                {
                    if (close)
                        e.Cancel = false;
                    else
                    {
                        e.Cancel = true;
                        comboBox2.Focus();
                        MessageBox.Show(E.Message, "предупреждение!");
                    }
                }
            }
            //else
            if (comboBox1.Visible)
                try
                {
                    if (comboBox1.SelectedIndex == -1 || comboBox1.Items == null)
                        throw new Exception("не выбрано основание для нового расписания");
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
