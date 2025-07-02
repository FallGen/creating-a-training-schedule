using System;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Windows.Forms;

namespace creating_a_training_schedule
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            update_list_week();             // обновление листа расписаний
            load_number_cloakroom();        // загрузка листа раздевалок
            load_list_teams();              // загрузка листа команд
            update_default_time();          // загрузка листа предустановленного времени
            load_settings();                // загрузка настроек
        }

        private void update_list_week()
        {
            try
            {
                lb_list_week.Items.Clear();
                var dir = new DirectoryInfo(path_save_week); // папка с файлами 

                int i = 0;
                foreach (FileInfo file in dir.GetFiles())
                {
                    i++;
                }

                string[] name_file = new string[i];

                int k = 0;
                foreach (FileInfo file in dir.GetFiles())
                {
                    name_file[k] = Path.GetFileName(file.FullName);
                    k++;
                }

                for (int c = name_file.Length - 1; c >= 0; c--)
                    lb_list_week.Items.Add(name_file[c]);
            }
            catch { }
        }

        string path_save_week = $@"C:\Users\{Environment.UserName}\AppData\Roaming\training_schedule\week\";
        string path_save_config = $@"C:\Users\{Environment.UserName}\AppData\Roaming\training_schedule\config\";
        string path_save_defaul_time = $@"C:\Users\{Environment.UserName}\AppData\Roaming\training_schedule\time\";
        string path_save_settings = $@"C:\Users\{Environment.UserName}\AppData\Roaming\training_schedule\settings";

        private void save_week()
        {
            try
            {
                string name_week;
                if (dataGridView1[0, 0].Value == null || dataGridView1[0, 18].Value == null)
                    name_week = "неопределенная неделя";
                else
                    name_week = dataGridView1[0, 0].Value.ToString() + "- " + dataGridView1[0, 18].Value.ToString();

                using (BinaryWriter file = new BinaryWriter(File.Open(path_save_week + name_week, FileMode.Create)))
                {
                    try
                    {
                        for (int i = 0; i < dataGridView1.ColumnCount; i++)
                            for (int j = 0; j < dataGridView1.RowCount; j++)
                                if (dataGridView1[i, j].Value == null)
                                {
                                    dataGridView1[i, j].Value = "";
                                    file.Write(dataGridView1[i, j].Value.ToString());
                                }
                                else
                                    file.Write(dataGridView1[i, j].Value.ToString());
                    }
                    catch { }
                }

                update_list_week();
                lbl_selected_week.Text = "выбранное расписание: " + name_week;
            }
            catch
            {
                DirectoryInfo qwe = new DirectoryInfo(path_save_week);
                qwe.Create();
                save_week();
            }
        }

        private string get_day(DateTime dateTime)
        {

            string day = "none";

            DateTime date = dateTime;
            int dayofweek;
            dayofweek = (int)date.DayOfWeek;

            switch (dayofweek)
            {
                case 0:
                    day = "ВС";
                    break;
                case 1:
                    day = "ПН";
                    break;
                case 2:
                    day = "ВТ";
                    break;
                case 3:
                    day = "СР";
                    break;
                case 4:
                    day = "ЧТ";
                    break;
                case 5:
                    day = "ПТ";
                    break;
                case 6:
                    day = "СБ";
                    break;
            }

            return day;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

            //using (var backGroundPen = new Pen(e.CellStyle.BackColor, 3))
            //using (var gridlinePen = new Pen(dataGridView1.GridColor, 3))



            //using (var selectedPen = new Pen(Color.Red, 8))
            //{
            //var bottomRightPoint = new Point(e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
            //var bottomleftPoint = new Point(e.CellBounds.Left, e.CellBounds.Bottom - 1);



            //    //e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);
            //    //e.Graphics.DrawRectangle(selectedPen, new Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width - 1, e.CellBounds.Height - 1));
            //    for (int i = 0; i < dataGridView1.RowCount; i += 3)
            //    {
            //        int x = this.dataGridView1.GetCellDisplayRectangle(1, 1, true).Bottom+10;
            //        int y = this.dataGridView1.GetCellDisplayRectangle(1, 1, true).Width;

            //        dataGridView1[0, i].Value = i;
            //        e.Graphics.DrawLine(selectedPen,  new Point(x), new Point(y));
            //        //e.Graphics.DrawLine(selectedPen,  new Point(300), new Point(300));

            //    }


            //}


            /*
            if (e.ColumnIndex > -1 & e.RowIndex > -1)
            {
                //Pen for left and top borders
                using (var backGroundPen = new Pen(e.CellStyle.BackColor, 1))
                //Pen for bottom and right borders
                using (var gridlinePen = new Pen(dataGridView1.GridColor, 1))
                //Pen for selected cell borders
                using (var selectedPen = new Pen(Color.Red, 1))
                {
                    var topLeftPoint = new Point(e.CellBounds.Left, e.CellBounds.Top);
                    var topRightPoint = new Point(e.CellBounds.Right - 1, e.CellBounds.Top);
                    var bottomRightPoint = new Point(e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                    var bottomleftPoint = new Point(e.CellBounds.Left, e.CellBounds.Bottom - 1);

                    //Draw selected cells here
                    if (this.dataGridView1[e.ColumnIndex, e.RowIndex].Selected)
                    {
                        //Paint all parts except borders.
                        e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);

                        //Draw selected cells border here
                        e.Graphics.DrawRectangle(selectedPen, new Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width - 1, e.CellBounds.Height - 1));

                        //Handled painting for this cell, Stop default rendering.
                        e.Handled = true;
                    }
                    //Draw non-selected cells here
                    //else
                    //{
                    //    //Paint all parts except borders.
                    //    e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);

                    //    //Top border of first row cells should be in background color
                    //    if (e.RowIndex == 0)
                    //        e.Graphics.DrawLine(backGroundPen, topLeftPoint, topRightPoint);

                    //    //Left border of first column cells should be in background color
                    //    if (e.ColumnIndex == 0)
                    //        e.Graphics.DrawLine(backGroundPen, topLeftPoint, bottomleftPoint);

                    //    //Bottom border of last row cells should be in gridLine color
                    //    if (e.RowIndex == dataGridView1.RowCount - 1)
                    //        e.Graphics.DrawLine(gridlinePen, bottomRightPoint, bottomleftPoint);
                    //    else  //Bottom border of non-last row cells should be in background color
                    //        e.Graphics.DrawLine(backGroundPen, bottomRightPoint, bottomleftPoint);

                    //    //Right border of last column cells should be in gridLine color
                    //    if (e.ColumnIndex == dataGridView1.ColumnCount - 1)
                    //        e.Graphics.DrawLine(gridlinePen, bottomRightPoint, topRightPoint);
                    //    else //Right border of non-last column cells should be in background color
                    //        e.Graphics.DrawLine(backGroundPen, bottomRightPoint, topRightPoint);

                    //    //Top border of non-first row cells should be in gridLine color, and they should be drawn here after right border
                    //    if (e.RowIndex > 0)
                    //        e.Graphics.DrawLine(gridlinePen, topLeftPoint, topRightPoint);

                    //    //Left border of non-first column cells should be in gridLine color, and they should be drawn here after bottom border
                    //    if (e.ColumnIndex > 0)
                    //        e.Graphics.DrawLine(gridlinePen, topLeftPoint, bottomleftPoint);

                    //    //We handled painting for this cell, Stop default rendering.
                    //    e.Handled = true;
                    //}
                }
            }

            */
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            //dataGridView1.Rows[e.RowIndex].Selected = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            save_week();
            MessageBox.Show("расписание сохранено", "успех!");
        }

        private void load_week(string week)
        {
            try
            {
                //update_table(true);
                using (BinaryReader file = new BinaryReader(File.Open(path_save_week + week, FileMode.Open)))
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                        for (int j = 0; j < dataGridView1.RowCount; j++)
                            dataGridView1[i, j].Value = file.ReadString();
                }
            }
            catch { }
        }

        private void update_table(bool new_table) // true - чистая таблица с новой датой | false - перерисовать с сохранением данных
        {
            if (new_table) dataGridView1.RowCount = 0;

            //int value_columns = dataGridView1.ColumnCount = Convert.ToInt32(tb_value_columns.Text);
            dataGridView1.RowCount = 21;

            for (int c = 1; c < 14; c++)
                dataGridView1.Columns[c].Width = 130;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 1; j < 14; j++)
                {
                    DataGridViewComboBoxCell qwe = new DataGridViewComboBoxCell();
                    qwe.DropDownWidth = 200;
                    //время
                    //if (i % 3 == 0)
                    //for (int q = 0; q < comboBox1.Items.Count; q++)
                    //    qwe.Items.Add(comboBox1.Items[q]);

                    //раздевалка
                    if (i % 3 == 1)
                        for (int q = 0; q < lb_number_cloakroom.Items.Count; q++)
                        {
                            qwe.Items.Add(lb_number_cloakroom.Items[q]);
                            qwe.ToolTipText = "номер раздевалки";
                        }

                    //команда
                    if (i % 3 == 2)
                        for (int q = 0; q < lb_teams.Items.Count; q++)
                        {
                            qwe.Items.Add(lb_teams.Items[q]);
                            qwe.ToolTipText = "команда";
                        }

                    if (i % 3 == 1 || i % 3 == 2)
                        dataGridView1.Rows[i].Cells[j] = qwe;
                }
            }

            dataGridView1.RowTemplate.Height = 70;
        }

        private void update_datetime(DateTime date)
        {
            int k = 0;
            for (int i = 0; i < dataGridView1.RowCount; i += 3)
            {
                DateTime temp_date = date.AddDays(k);

                dataGridView1[0, i].Value = temp_date.ToShortDateString() + " " + get_day(temp_date) + " ";
                //dataGridView1[0, i + 1].Value = get_day(temp_date);
                k++;
            }
        }

        private void button8_Click(object sender, EventArgs e) // создать расписание
        {
            try
            {
                Form3 form = new Form3();

                // на основе расписания
                for (int q = 0; q < lb_list_week.Items.Count; q++)
                    form.comboBox1.Items.Add(lb_list_week.Items[q].ToString());

                try
                {
                    var dir = new DirectoryInfo(path_save_defaul_time); // папка с файлами 

                    int i = 0;
                    foreach (FileInfo file in dir.GetFiles())
                    {
                        i++;
                    }

                    string[] name_file = new string[i];

                    int k = 0;
                    foreach (FileInfo file in dir.GetFiles())
                    {
                        name_file[k] = Path.GetFileName(file.FullName);
                        k++;
                    }

                    for (int c = name_file.Length - 1; c >= 0; c--)
                        form.comboBox2.Items.Add(name_file[c]);
                }
                catch { }


                if (DialogResult.OK == form.ShowDialog())
                {
                    if (form.radioButton1.Checked) // чистое расписание
                        update_table(true);

                    if (form.radioButton2.Checked) // на основе существующего расписания
                        load_week(form.comboBox1.SelectedItem.ToString());

                    if (form.checkBox1.Checked)
                    {
                        //update_table(false);
                        string[] data_time = new string[dataGridView1.ColumnCount];

                        using (BinaryReader file = new BinaryReader(File.Open(path_save_defaul_time + form.comboBox2.SelectedItem.ToString(), FileMode.Open)))
                        {
                            for (int c = 0; c < data_time.Length; c++)
                                data_time[c] = file.ReadString();
                        }

                        for (int c = 0; c < dataGridView1.RowCount; c += 3)
                        {
                            int count = 0;
                            for (int v = 1; v < dataGridView1.ColumnCount; v++)
                            {
                                dataGridView1[v, c].Value = data_time[count];
                                count++;
                            }
                        }
                    }

                    update_datetime(form.dateTimePicker1.Value);
                    save_week();
                }
            }
            catch { }
        }

        private void reload_table()
        {
            try
            {
                string[,] str = new string[dataGridView1.ColumnCount, dataGridView1.RowCount];

                for (int i = 0; i < str.GetLength(0); i++)
                    for (int j = 0; j < str.GetLength(1); j++)
                        if (dataGridView1[i, j].Value == null)
                            str[i, j] = "";
                        else
                            str[i, j] = dataGridView1[i, j].Value.ToString();

                update_table(true);

                for (int i = 0; i < str.GetLength(1); i++)
                    for (int j = 0; j < str.GetLength(0); j++)
                        dataGridView1[j, i].Value = str[j, i].ToString();
            }
            catch { }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            form.Text = "добавление новой раздевалки";
            form.label1.Text = "номер раздевалки:";

            if (DialogResult.OK == form.ShowDialog())
            {
                lb_number_cloakroom.Items.Add(form.textBox1.Text);
                reload_table();
                save_number_cloakroom();
                MessageBox.Show("раздевалка добавлена", "успех!");
            }
        }
        private void save_number_cloakroom()
        {
            try
            {
                using (BinaryWriter file = new BinaryWriter(File.Open(path_save_config + "list_number_cloakroom", FileMode.Create)))
                {
                    file.Write(lb_number_cloakroom.Items.Count);
                    for (int i = 0; i < lb_number_cloakroom.Items.Count; i++)
                        file.Write(lb_number_cloakroom.Items[i].ToString());
                }
            }
            catch
            {
                DirectoryInfo qwe = new DirectoryInfo(path_save_config);
                qwe.Create();
                save_number_cloakroom();
            }
        }

        private void load_number_cloakroom()
        {
            try
            {
                using (BinaryReader file = new BinaryReader(File.Open(path_save_config + "list_number_cloakroom", FileMode.Open)))
                {
                    int lenght = file.ReadInt32();
                    for (int i = 0; i < lenght; i++)
                        lb_number_cloakroom.Items.Add(file.ReadString());
                }
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            form.Text = "добавление новой команды";
            form.label1.Text = "наименование команды:";

            if (DialogResult.OK == form.ShowDialog())
            {
                lb_teams.Items.Add(form.textBox1.Text);
                reload_table();
                save_list_teams();
                MessageBox.Show("команда добавлена", "успех!");
            }
        }
        private void save_list_teams()
        {
            try
            {
                using (BinaryWriter file = new BinaryWriter(File.Open(path_save_config + "list_teams", FileMode.Create)))
                {
                    file.Write(lb_teams.Items.Count);
                    for (int i = 0; i < lb_teams.Items.Count; i++)
                        file.Write(lb_teams.Items[i].ToString());
                }
            }
            catch
            {
                DirectoryInfo qwe = new DirectoryInfo(path_save_config);
                qwe.Create();
                save_list_teams();
            }
        }
        private void load_list_teams()
        {
            try
            {
                using (BinaryReader file = new BinaryReader(File.Open(path_save_config + "list_teams", FileMode.Open)))
                {
                    int lenght = file.ReadInt32();
                    for (int i = 0; i < lenght; i++)
                        lb_teams.Items.Add(file.ReadString());
                }
            }
            catch { }
        }

        private void button10_Click(object sender, EventArgs e) // добавление нового присета времени
        {
            //if (lb_default_time.Items.Count > 0 && lb_default_time.SelectedItems != null)
            open_default_time(false);
        }

        private void open_default_time(bool flag) // true - изменение | false - добавление 
        {
            try
            {
                //if (lb_default_time.Items.Count > 0 && lb_default_time.SelectedItems != null)
                {

                    Form4 form = new Form4();

                    if (flag)
                    {
                        form.Text = "изменение присета времени";

                        string name_time = lb_default_time.SelectedItem.ToString();
                        string[] data_time = new string[form.dataGridView1.RowCount];

                        using (BinaryReader file = new BinaryReader(File.Open(path_save_defaul_time + name_time, FileMode.Open)))
                        {
                            for (int i = 0; i < data_time.Length; i++)
                                data_time[i] = file.ReadString();
                        }

                        form.textBox1.Text = name_time;
                        for (int i = 0; i < form.dataGridView1.RowCount; i++)
                            form.dataGridView1[0, i].Value = data_time[i];
                    }
                    else
                        form.Text = "добавление присета времени";

                    if (DialogResult.OK == form.ShowDialog())
                    {
                        string name_time = form.textBox1.Text;
                        string[] data_time = new string[form.dataGridView1.RowCount];

                        for (int i = 0; i < form.dataGridView1.RowCount; i++)
                            if (form.dataGridView1[0, i].Value == null)
                                data_time[i] = "";
                            else
                                data_time[i] = form.dataGridView1[0, i].Value.ToString();

                        //lb_default_time.Items.Add(name_time);
                        save_default_time(data_time, name_time);
                        update_default_time();
                    }
                }
            }
            catch { }
        }

        private void save_default_time(string[] data_time, string name_time)
        {
            try
            {
                using (BinaryWriter file = new BinaryWriter(File.Open(path_save_defaul_time + name_time, FileMode.Create)))
                {
                    for (int i = 0; i < data_time.Length; i++)
                        file.Write(data_time[i].ToString());
                }
            }
            catch
            {
                DirectoryInfo qwe = new DirectoryInfo(path_save_defaul_time);
                qwe.Create();
                save_default_time(data_time, name_time);
            }
        }

        private void update_default_time()
        {
            try
            {
                lb_default_time.Items.Clear();
                var dir = new DirectoryInfo(path_save_defaul_time); // папка с файлами 

                int i = 0;
                foreach (FileInfo file in dir.GetFiles())
                {
                    i++;
                }

                string[] name_file = new string[i];

                int k = 0;
                foreach (FileInfo file in dir.GetFiles())
                {
                    name_file[k] = Path.GetFileName(file.FullName);
                    k++;
                }

                for (int c = name_file.Length - 1; c >= 0; c--)
                    lb_default_time.Items.Add(name_file[c]);
            }
            catch { }
        }

        private void button11_Click(object sender, EventArgs e) // изменение присета времени
        {
            if (lb_default_time.Items.Count > 0 && lb_default_time.SelectedItems != null)
                open_default_time(true);
            else
            {
                MessageBox.Show("не выбрано время пресета по умолчанию для изменения", "предупреждение");
            }
        }

        private void ll_change_time_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) // изменить всё время в расписании
        {
            try
            {
                if (lb_list_week.SelectedIndex != -1)
                {

                    Form5 form = new Form5();
                    try
                    {
                        var dir = new DirectoryInfo(path_save_defaul_time); // папка с файлами 

                        int i = 0;
                        foreach (FileInfo file in dir.GetFiles())
                        {
                            i++;
                        }

                        string[] name_file = new string[i];

                        int k = 0;
                        foreach (FileInfo file in dir.GetFiles())
                        {
                            name_file[k] = Path.GetFileName(file.FullName);
                            k++;
                        }

                        for (int c = name_file.Length - 1; c >= 0; c--)
                            form.comboBox1.Items.Add(name_file[c]);
                    }
                    catch { }

                    if (DialogResult.OK == form.ShowDialog())
                    {
                        if (dataGridView1.RowCount > 0)
                        {
                            string[] data_time = new string[dataGridView1.ColumnCount - 1];

                            using (BinaryReader file = new BinaryReader(File.Open(path_save_defaul_time + form.comboBox1.SelectedItem.ToString(), FileMode.Open)))
                            {
                                for (int c = 0; c < data_time.Length; c++)
                                    data_time[c] = file.ReadString();
                            }

                            for (int c = 0; c < dataGridView1.RowCount; c += 3)
                            {
                                int count = 0;
                                for (int v = 1; v < dataGridView1.ColumnCount; v++)
                                {
                                    dataGridView1[v, c].Value = data_time[count];
                                    count++;
                                }
                            }
                        }
                        else { MessageBox.Show("не выбрано расписание", "предупреждение!"); }
                    }
                }
                else MessageBox.Show("не выбрано расписание", "предупреждение!");
            }
            catch { }
        }

        private void button7_Click(object sender, EventArgs e) // удаление расписания
        {
            try
            {
                if (lb_list_week.SelectedIndex != -1)
                {
                    if (DialogResult.Yes == MessageBox.Show("удалить расписание " + lb_list_week.SelectedItem.ToString() + "?", "предупреждение", MessageBoxButtons.YesNo))
                    {
                        FileInfo fileInf = new FileInfo(path_save_week + lb_list_week.SelectedItem.ToString());
                        fileInf.Delete();

                        update_list_week();
                        lb_list_week.SelectedIndex = -1;
                        dataGridView1.RowCount = 0;
                        lbl_selected_week.Text = "выбранное расписание: -";
                    }
                }
                else MessageBox.Show("не выбрано расписание", "предупреждение");
            }
            catch { }
        }

        private void button4_Click(object sender, EventArgs e)  // удаление раздевалки
        {
            try
            {
                if (lb_number_cloakroom.SelectedIndex != -1)
                {
                    if (DialogResult.Yes == MessageBox.Show("удалить раздевалку " + lb_number_cloakroom.SelectedItem.ToString() + "?", "предупреждение", MessageBoxButtons.YesNo))
                    {
                        lb_number_cloakroom.Items.RemoveAt(lb_number_cloakroom.SelectedIndex);
                        reload_table();
                        save_number_cloakroom();
                        lb_number_cloakroom.SelectedIndex = -1;
                    }
                }
                else MessageBox.Show("не выбрана раздевалка", "предупреждение");
            }
            catch { }
        }

        private void button5_Click(object sender, EventArgs e) // удаление команды
        {
            try
            {
                if (lb_teams.SelectedIndex != -1)
                {
                    if (DialogResult.Yes == MessageBox.Show("удалить команду " + lb_teams.SelectedItem.ToString() + "?", "предупреждение", MessageBoxButtons.YesNo))
                    {
                        lb_teams.Items.RemoveAt(lb_teams.SelectedIndex);
                        reload_table();
                        save_list_teams();
                        lb_teams.SelectedIndex = -1;
                    }
                }
                else MessageBox.Show("не выбрана команда", "предупреждение");
            }
            catch { }
        }

        private void button9_Click(object sender, EventArgs e) // удаление времени по умолчанию
        {
            try
            {
                if (lb_default_time.SelectedIndex != -1)
                {
                    if (DialogResult.Yes == MessageBox.Show("удалить время по умолчанию " + lb_default_time.SelectedItem.ToString() + "?", "предупреждение", MessageBoxButtons.YesNo))
                    {
                        FileInfo fileInf = new FileInfo(path_save_defaul_time + lb_default_time.SelectedItem.ToString());
                        fileInf.Delete();

                        update_default_time();
                        lb_default_time.SelectedIndex = -1;
                    }
                }
                else MessageBox.Show("не выбрано время по умолчанию", "предупреждение");

            }
            catch { }

        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            update_printer();
        }

        private void update_printer()
        {
            try
            {
                CB_have_printer.Items.Clear();
                foreach (string printerName in PrinterSettings.InstalledPrinters)
                {
                    CB_have_printer.Items.Add(printerName);
                }
            }
            catch { }
        }

        private void CB_have_printer_SelectedIndexChanged(object sender, EventArgs e) //обновление комбобокса с принтерами
        {
            save_settings();
        }

        private void save_settings()
        {
            try
            {
                using (BinaryWriter file = new BinaryWriter(File.Open(path_save_settings, FileMode.Create)))
                {
                    if (CB_have_printer.SelectedItem != null)
                        file.Write(CB_have_printer.SelectedItem.ToString());
                    else file.Write("empty");
                }
            }
            catch
            {
                save_settings();
            }
        }

        private void load_settings()
        {
            try
            {
                using (BinaryReader file = new BinaryReader(File.Open(path_save_settings, FileMode.Open)))
                {
                    CB_have_printer.Items.Add(file.ReadString().ToString());
                }
                CB_have_printer.SelectedIndex = 0;
            }
            catch { }
        }

        private void button6_Click(object sender, EventArgs e) // печать
        {
            try
            {
                if (lb_list_week.SelectedIndex != -1)
                {
                    if (CB_have_printer.SelectedIndex != -1)
                    {
                        create_excel_doc("print");
                        //создать метод для печать + подключить библиотеку
                    }
                    else
                        MessageBox.Show("не выбран принтер", "предупреждение");
                }
                else MessageBox.Show("не выбрано расписание для печати", "предупреждение");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void despose_excel()
        {
            try
            {
                Process[] localByName = Process.GetProcessesByName("Excel");
                foreach (Process p in localByName)
                {
                    p.Kill();
                }
            }
            catch { }

        }
        private void create_excel_doc(string flag) // exprot | print
        {
            despose_excel();
            string path_excel = $@"C:\Users\{Environment.UserName}\AppData\Roaming\training_schedule\print.xlsx";

            using (Excel_improt.Excel_improt excel = new Excel_improt.Excel_improt())
            {
                //удалить существующий эксель файл
                excel.delete(path_excel);

                //загрузка данных в эксель
                if (excel.Open(filePath: Path.Combine(Environment.CurrentDirectory, path_excel)))
                {
                    int k = 0;
                    for (char c = 'B'; c <= 'O'; c++)
                    {
                        int j = 0;
                        for (int i = 3; i < dataGridView1.RowCount + 3; i++)
                        {
                            excel.Set(column: c.ToString(), row: i, data: dataGridView1[k, j].Value.ToString());
                            j++;
                        }
                        k++;
                    }
                }

                //format_cells - форматирование
                string name_week = dataGridView1[0, 0].Value.ToString() + "- " + dataGridView1[0, 18].Value.ToString();

                excel.format_cells("Расписание тренировок " + name_week);
                excel.Save();

                if (flag == "export")
                {
                    despose_excel();
                    Process.Start(path_excel);
                }
                else
                    if (flag == "print")
                {
                    excel.print_doc(CB_have_printer.SelectedItem.ToString());
                    despose_excel();
                }
            }


        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (lb_list_week.SelectedIndex != -1)
                {
                    create_excel_doc("export");

                }
                else MessageBox.Show("не выбрано расписание для экспорта", "предупреждение");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void btn1_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 0; c < 3; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";
        }

        private void lb_list_week_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (/*lb_list_week.SelectedItems[lb_list_week.SelectedIndex] != null && */lb_list_week.SelectedIndex > -1)
                {
                    update_table(true);
                    load_week(lb_list_week.SelectedItems[lb_list_week.SelectedIndex].ToString());
                    lbl_selected_week.Text = "выбранное расписание: " + lb_list_week.SelectedItems[lb_list_week.SelectedIndex].ToString();
                }
            }
            catch { }
        }

        private void btn2_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 3; c < 6; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";
        }

        private void btn3_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 6; c < 9; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";
        }

        private void btn4_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 9; c < 12; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";

        }

        private void btn5_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 12; c < 15; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";

        }

        private void btn6_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 15; c < 18; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";

        }

        private void btn7_del_1day_Click(object sender, EventArgs e)
        {
            for (int c = 18; c < 21; c++)
                for (int i = 1; i < dataGridView1.ColumnCount; i++)
                    dataGridView1[i, c].Value = "-";

        }

        private void lb_list_week_SelectedValueChanged(object sender, EventArgs e)
        {
            if (lb_list_week.SelectedItems.Count > 0 && lb_list_week.SelectedItems != null)
            {
                update_table(true);
                load_week(lb_list_week.SelectedItem.ToString());
                lbl_selected_week.Text = "выбранное расписание: " + lb_list_week.SelectedItem.ToString();
            }
        }
    }
}