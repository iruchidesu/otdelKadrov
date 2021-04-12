using System;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;

namespace WindowsFormsApplication1
{
    public partial class Student : Form
    {
        public Student()
        {
            InitializeComponent();
        }

        ArrayList districts = new ArrayList();
        ArrayList groups = new ArrayList();
        ClassPrintForm4 printForm4 = new ClassPrintForm4();
        PrintDistrict printDistrict = new PrintDistrict();
        PrintGroup PrintGroup = new PrintGroup();
        PrintForm3 printForm3 = new PrintForm3();
        PrintForm8 printForm8 = new PrintForm8();
        PrintForm9 printForm9 = new PrintForm9();
        PrintForm10 printForm10 = new PrintForm10();

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void LoadTable()
        {
            string select = @"SELECT     Student.name, Student.birth, sex.sex, Goden.Goden, Student.katGodnost, Student.indx, city.city, district.district, Student.country, Student.street, 
                                  Student.house, Student.flat, Student.phone, Student.homePhone,  Student.ciizenship, Student.passpSeries, Student.passpNumber, Student.passpKemVidan, Student.passpDate, 
                                  [Group].groupName, otdelenie.number, comm.type, Student.prikazNumIn, Student.dateIn, Student.prikazNumOut, Student.dateOut, Student.prichinaOut, 
                                  Student.kval, Student.prikazNumKval, Akadem_otpusk.value as academ, Student.note,  Student.id
                              FROM         [Group] INNER JOIN
                                  Student ON [Group].id = Student.idGroup INNER JOIN
                                  sex ON Student.id_sex = sex.id INNER JOIN
                                  otdelenie ON Student.id_otdelenie = otdelenie.id INNER JOIN
                                  Goden ON Student.id_goden = Goden.id INNER JOIN
                                  district ON Student.id_district = district.id INNER JOIN
                                  comm ON Student.id_comm = comm.id INNER JOIN
                                  city ON Student.id_city = city.id INNER JOIN
                                  Akadem_otpusk ON Student.id_academ = Akadem_otpusk.id "; 

            if (обучающиесяToolStripMenuItem.Checked == true)
            {
                select += @"WHERE (prikazNumKval = '') AND (prikazNumOut = '')";
                label10.Text = "Всего обучающихся студентов: ";
            }

           

            DataSet ds1 = new DataSet();
            DataSet ds2 = new DataSet();
            try
            {
                ds1 = Util.FillTable("Student", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            string select1;

            select1 = @"SELECT Goden FROM Goden";
            try
            {
                ds2 = Util.FillTable("Goden", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                Goden.Items.Add(row.ItemArray[0]);
            }

            select1 = @"SELECT groupName FROM [Group]";
            try
            {
                ds2 = Util.FillTable("Group", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                groupName.Items.Add(row.ItemArray[0]);
                groups.Add(row.ItemArray[0]);
            }

            select1 = @"SELECT city FROM city";
            
            try
            {
                ds2 = Util.FillTable("city", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                city.Items.Add(row.ItemArray[0]);
            }
            
            select1 = @"SELECT district FROM district";
            try
            {
                ds2 = Util.FillTable("district", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                district.Items.Add(row.ItemArray[0]);
                districts.Add(row.ItemArray[0]);
            }

            select1 = @"SELECT number FROM otdelenie";
            try
            {
                ds2 = Util.FillTable("otdelenie", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                otdelenie.Items.Add(row.ItemArray[0]);
            }

            select1 = @"SELECT type FROM comm";
            try
            {
                ds2 = Util.FillTable("comm", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                comm.Items.Add(row.ItemArray[0]);
            }

            select1 = @"SELECT value FROM Akadem_otpusk";
            try
            {
                ds2 = Util.FillTable("value", select1);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            foreach (DataRow row in ds2.Tables[0].Rows)
            {
                value.Items.Add(row.ItemArray[0]);
            }

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = ds1;
            dataGridView1.DataMember = "Student";

            var count = dataGridView1.Rows.Count - 1;
            label10.Text += count.ToString();

            checkBox3.Text = "В академическом\n отпуске";

            foreach (string row in groups)
            {
                comboBox4.Items.Add(row);
            }

            foreach (string row in districts)
            {
                comboBox5.Items.Add(row);
            }
        }

        public void LoadTable2()
        {            
            string selectAdd = "";

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            string select = @"SELECT Student.name, Student.birth, sex.sex, Goden.Goden, Student.katGodnost, Student.indx, city.city, district.district, Student.country, Student.street, 
                                  Student.house, Student.flat, Student.phone, Student.homePhone, Student.ciizenship, Student.passpSeries, Student.passpNumber, Student.passpKemVidan, Student.passpDate, 
                                  [Group].groupName, otdelenie.number, comm.type, Student.prikazNumIn, Student.dateIn, Student.prikazNumOut, Student.dateOut, Student.prichinaOut, 
                                  Student.kval, Student.prikazNumKval, Akadem_otpusk.value as academ, Student.note, Student.id
                              FROM [Group] INNER JOIN
                                  Student ON [Group].id = Student.idGroup INNER JOIN
                                  sex ON Student.id_sex = sex.id INNER JOIN
                                  otdelenie ON Student.id_otdelenie = otdelenie.id INNER JOIN
                                  Goden ON Student.id_goden = Goden.id INNER JOIN
                                  district ON Student.id_district = district.id INNER JOIN
                                  comm ON Student.id_comm = comm.id INNER JOIN
                                  city ON Student.id_city = city.id INNER JOIN
                                  Akadem_otpusk ON Student.id_academ = Akadem_otpusk.id ";
            if (обучающиесяToolStripMenuItem.Checked == true)
            {
                selectAdd = @"WHERE (prikazNumKval = '') AND (prikazNumOut = '')";
                label10.Text = "Всего обучающихся студентов: ";
                checkBox2.Enabled = false;
                otchislennieDateTimePicker3.Enabled = false;
                otchislennieDateTimePicker4.Enabled = false;
            }
            else if (выпускникиToolStripMenuItem.Checked == true)
            {
                selectAdd = @"WHERE (prikazNumKval != '') AND (prikazNumOut = '')";
                label10.Text = "Всего выпускников: ";
                checkBox2.Enabled = false;
                otchislennieDateTimePicker3.Enabled = false;
                otchislennieDateTimePicker4.Enabled = false;
            }
            else if (отчисленныеToolStripMenuItem.Checked == true)
            {
                selectAdd = @"WHERE ((prikazNumKval != '') OR (prikazNumOut != ''))";
                label10.Text = "Всего отчисленных: ";
                checkBox2.Enabled = false;
                otchislennieDateTimePicker3.Enabled = false;
                otchislennieDateTimePicker4.Enabled = false;
            }
            else if (толькоОтчисленныеToolStripMenuItem.Checked == true)
            {
                selectAdd = @"WHERE (prikazNumKval = '') AND (prikazNumOut != '')";
                label10.Text = "Всего отчисленных: ";
                checkBox2.Enabled = true;
                otchislennieDateTimePicker3.Enabled = true;
                otchislennieDateTimePicker4.Enabled = true;
            }


            if (textBox1.Text != "")
                selectAdd = @" WHERE name LIKE '" + textBox1.Text + "%' ";

            select += selectAdd;

            if (comboBox4.Text != "")
                select += @" AND groupName = '" + comboBox4.SelectedItem + "'";

            if (comboBox5.Text != "")            
                select += @" AND id_district = '" + ConvertDistrictNameToIdDistrict(comboBox5.Text) + "'";

            if (comboBox1.Text != "")
                select += @" AND id_sex = '" + ConvertSexValueToIdSex(comboBox1.Text) + "'";
            else
                select += " ";

            if (comboBox3.Text != "")
                select += @" AND id_otdelenie = '" + ConvertOtdelenieNumberToIdOtdelenie(comboBox3.Text) + "'";

            if (comboBox2.Text != "")
                select += @" AND id_comm = '" + ConvertCommTypeToIdComm(comboBox2.Text) + "'";

            if (checkBox1.Checked == true)
                select += @" AND (birth >= '" + birthDateTimePicker1.Value + "' AND birth <= '" + birthDateTimePicker2.Value + "')";

            if (checkBox2.Checked == true)
                select += @" AND (dateOut >= '" + otchislennieDateTimePicker3.Value + "' AND dateOut <= '" + otchislennieDateTimePicker4.Value + "')";
            
            if (checkBox3.Checked == true)
                select += @" AND id_academ != 2";

            DataSet ds1 = new DataSet();
            try
            {
                ds1 = Util.FillTable("Student", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = ds1;
            dataGridView1.DataMember = "Student";

            var count = dataGridView1.Rows.Count - 1;
            label10.Text += count.ToString();

            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void Student_Load(object sender, EventArgs e)
        {
            LoadTable();
        }

        int ConvertGroupNameToIdGroup(string groupName) //определение ид группы из таблицы Group
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM [Group] WHERE ( groupName = '" + groupName + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        int ConvertAcademValueToIdAcadem(string value) //определение ид для академа из таблицы akadem_otpusk
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Akadem_otpusk", "SELECT id FROM Akadem_otpusk WHERE (value = '" + value + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return int.Parse(ds.Tables[0].Rows[0].ItemArray[0].ToString());
        }

        int ConvertSexValueToIdSex(string value) //определение ид пола из таблицы sex
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM sex WHERE ( sex = '" + value + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return int.Parse(ds.Tables[0].Rows[0].ItemArray[0].ToString());
        }

        int ConvertCityNameToIdCity(string city)
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM city WHERE ( city = '" + city + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        int ConvertDistrictNameToIdDistrict(string districtName)
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM district WHERE ( district = '" + districtName + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        int ConvertGodenToIdGoden(string goden) //определение ид ВО из таблицы Goden
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM Goden WHERE ( Goden = '" + goden + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return int.Parse(ds.Tables[0].Rows[0].ItemArray[0].ToString());
        }

        int ConvertOtdelenieNumberToIdOtdelenie(string otdelenieNumber) //определение ид отделения из таблицы otdelenie
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM otdelenie WHERE ( number = '" + otdelenieNumber + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return int.Parse(ds.Tables[0].Rows[0].ItemArray[0].ToString());
        }

        int ConvertCommTypeToIdComm(string type) //определение ид comm из таблицы comm
        {
            DataSet ds;
            try
            {
                ds = Util.FillTable("Student", "SELECT id FROM comm WHERE ( type = '" + type + "')");
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return 0;
            }
            return int.Parse(ds.Tables[0].Rows[0].ItemArray[0].ToString());
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataSet ds1;
            if (dataGridView1.CurrentRow.Cells[2].Value.ToString() == "")
            {
                try
                {
                    if (String.IsNullOrWhiteSpace(dataGridView1.CurrentRow.Cells[0].Value.ToString()))
                        MessageBox.Show("Необходимо заполнить имя");
                    else
                    {
                        string select = "INSERT INTO Student (name, homePhone, prikazNumKval, prikazNumOut) VALUES ('" + dataGridView1.CurrentRow.Cells[0].Value + "', '', '', '')";
                        Util.FillTable("Student", select);
                        string select2 = @"SELECT id FROM Student WHERE name = '" + dataGridView1.CurrentRow.Cells[0].Value + "' AND id_sex = 5 ";
                        ds1 = Util.FillTable("Student", select2);
                        dataGridView1.CurrentRow.Cells[31].Value = ds1.Tables[0].Rows[0].ItemArray[0].ToString();
                        AddStudentSex add_stud = new AddStudentSex();
                        add_stud.Location = new System.Drawing.Point(MousePosition.X - 60, MousePosition.Y - 40);
                        add_stud.ShowDialog();
                        dataGridView1.CurrentRow.Cells[2].Value = add_stud.ChooseSex();
                        string select3 = @"UPDATE Student SET id_sex = " + ConvertSexValueToIdSex(dataGridView1.CurrentRow.Cells[2].Value.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                        Util.FillTable("upd_sex", select3);
                    }
                }
                catch
                {
                    MessageBox.Show("Необходимо заполнить имя");
                    return;
                }
            }
            else
            {
                string select = "UPDATE Student SET";
                switch (e.ColumnIndex)
                {
                    case 0:
                        select += " name ";
                        break;
                    case 1:
                        select += " birth ";
                        break;
                    case 2: //sex
                        return;
                    case 3: //goden
                        return;
                    case 4:
                        select += " katGodnost ";
                        break;
                    case 5:
                        select += " indx ";
                        break;
                    case 6: //city
                        return;
                    case 7: //district
                        return;
                    case 8:
                        select += " country ";
                        break;
                    case 9:
                        select += " street ";
                        break;
                    case 10:
                        select += " house ";
                        break;
                    case 11:
                        select += " flat ";
                        break;
                    case 12:
                        select += " phone ";
                        break;
                    case 13:
                        select += " homePhone ";
                        break;
                    case 14:
                        select += " ciizenship ";
                        break;
                    case 15:
                        select += " passpSeries ";
                        break;
                    case 16:
                        select += " passpNumber ";
                        break;
                    case 17:
                        select += " passpKemVidan ";
                        break;
                    case 18:
                        select += " passpDate ";
                        break;
                    case 19:
                        select += " idGroup = '" + ConvertGroupNameToIdGroup(dataGridView1.CurrentCell.Value.ToString()) + "' WHERE (id = '" + dataGridView1.CurrentRow.Cells[31].Value + "')";
                        try
                        {
                            Util.FillTable("Group", select);
                        }
                        catch (SqlException ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        return;
                    case 20: //otdelenie
                        return;
                    case 21: //comm
                        return;
                    case 22:
                        select += " prikazNumIn ";
                        break;
                    case 23:
                        select += " dateIn ";
                        break;
                    case 24:
                        select += " prikazNumOut ";
                        break;
                    case 25:
                        select += " dateOut ";
                        break;
                    case 26:
                        select += " prichinaOut ";
                        break;
                    case 27:
                        select += " kval ";
                        break;
                    case 28:
                        select += " prikazNumKval ";
                        break;
                    case 29: //academ_otpusk
                        return;
                    case 30:
                        select += " note ";
                        break;
                }
                select += " = '" + dataGridView1.CurrentCell.Value + "' WHERE (id = '" + dataGridView1.CurrentRow.Cells[31].Value + "')";

                try
                {
                    Util.FillTable("Student", select);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string insert;
            string update;
            if (e.ColumnIndex == city.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow && !String.IsNullOrWhiteSpace(dataGridView1.CurrentRow.Cells[0].Value.ToString()))
            {
                if (!this.city.Items.Contains(e.FormattedValue))
                {
                    this.city.Items.Add(e.FormattedValue);
                    insert = @"INSERT INTO city (city) VALUES ('" + e.FormattedValue + "')";
                    try
                    {
                        Util.FillTable("add_city", insert);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    update = @"UPDATE Student SET id_city = " + ConvertCityNameToIdCity(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                    try
                    {
                        Util.FillTable("upd_city", update);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
                }
                else
                {
                    update = @"UPDATE Student SET id_city = " + ConvertCityNameToIdCity(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                    try
                    {
                        Util.FillTable("upd_city", update);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
                }
            }

            if (e.ColumnIndex == district.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow)
            {
                if (!this.district.Items.Contains(e.FormattedValue))
                {
                    this.district.Items.Add(e.FormattedValue);
                    insert = @"INSERT INTO district (district) VALUES ('" + e.FormattedValue + "')";
                    try
                    {
                        Util.FillTable("add_district", insert);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    update = @"UPDATE Student SET id_district = " + ConvertDistrictNameToIdDistrict(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                    try
                    {
                        Util.FillTable("upd_district", update);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
                }
                else
                {
                    update = @"UPDATE Student SET id_district = " + ConvertDistrictNameToIdDistrict(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                    try
                    {
                        Util.FillTable("upd_district", update);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                    dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
                }
            }
            if (e.ColumnIndex == sex.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow)
            {
                update = @"UPDATE Student SET id_sex = " + ConvertSexValueToIdSex(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                try
                {
                    Util.FillTable("upd_sex", update);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
            if (e.ColumnIndex == Goden.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow)
            {
                update = @"UPDATE Student SET id_Goden = " + ConvertGodenToIdGoden(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                try
                {
                    Util.FillTable("upd_Goden", update);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
            if (e.ColumnIndex == otdelenie.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow)
            {
                update = @"UPDATE Student SET id_otdelenie = " + ConvertOtdelenieNumberToIdOtdelenie(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                try
                {
                    Util.FillTable("upd_otdelenie", update);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
            if (e.ColumnIndex == comm.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow)
            {
                update = @"UPDATE Student SET id_comm = " + ConvertCommTypeToIdComm(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                try
                {
                    Util.FillTable("upd_comm", update);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
            if (e.ColumnIndex == value.DisplayIndex && !dataGridView1.CurrentRow.IsNewRow)
            {
                update = @"UPDATE Student SET id_academ = " + ConvertAcademValueToIdAcadem(e.FormattedValue.ToString()) + " WHERE (id = " + dataGridView1.CurrentRow.Cells[31].Value + ")";
                try
                {
                    Util.FillTable("upd_academ", update);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                dataGridView1.CurrentRow.Cells[e.ColumnIndex].Value = e.FormattedValue;
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (this.dataGridView1.CurrentCellAddress.X == city.DisplayIndex)
            {
                ComboBox cb = e.Control as ComboBox;
                if (cb != null)
                {
                    cb.DropDownStyle = ComboBoxStyle.DropDown;
                }
            }

            if (this.dataGridView1.CurrentCellAddress.X == district.DisplayIndex)
            {
                ComboBox cb = e.Control as ComboBox;
                if (cb != null)
                {
                    cb.DropDownStyle = ComboBoxStyle.DropDown;
                }
            }
        }

        private void редакторСтрокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringEditor se = new StringEditor();
            se.ShowDialog();
            редакторСтрокToolStripMenuItem.CheckState = CheckState.Unchecked;
        }

        private void SetToolStripItemCheckedState(ToolStripItem itm)
        {
            foreach (ToolStripMenuItem menuItem in управлениеToolStripMenuItem.DropDownItems)
            {
                if (menuItem == itm)
                {
                    menuItem.CheckState = CheckState.Checked;
                }
                else
                {
                    menuItem.CheckState = CheckState.Unchecked;
                }
            }
        }

        private void управлениеToolStripMenuItem_DropDownItemClicked_1(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem != редакторСтрокToolStripMenuItem && e.ClickedItem != печатьToolStripMenuItem && e.ClickedItem != выходToolStripMenuItem)
                SetToolStripItemCheckedState(e.ClickedItem);
            if (e.ClickedItem == редакторСтрокToolStripMenuItem || e.ClickedItem == печатьToolStripMenuItem || e.ClickedItem == выходToolStripMenuItem)
                return;
            else
                LoadTable2();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
                LoadTable2();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
                LoadTable2();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
                LoadTable2();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
                LoadTable2();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
                LoadTable2();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            LoadTable2();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            LoadTable2();
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
                LoadTable2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            LoadTable2();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadTable2();
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            LoadTable2();
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            печатьToolStripMenuItem.CheckState = CheckState.Unchecked;
        }
        
        private void форма1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintGroup.ShowDialog();
        }

        private void форма2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDistrict.ShowDialog();
        }

        private void форма3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printForm3.ShowDialog();
        }
        
        private void форма4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            backgroundWorker1.RunWorkerAsync();
        }

        private void форма5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            backgroundWorker2.RunWorkerAsync();
        }

        private void форма8ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printForm8.ShowDialog();
        }

        private void форма9ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printForm9.ShowDialog();
        }

        private void форма10ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printForm10.ShowDialog();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            printForm4.PrintForm4();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            printForm4.PrintForm5();
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void comboBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadTable2();
            }
        }

        private void comboBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadTable2();
            }
        }

        private void comboBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadTable2();
            }
        }

        private void comboBox5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadTable2();
            }
        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LoadTable2();
            }
        }

        private void печатьВыборкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDataGrid print = new PrintDataGrid(dataGridView1.Rows);
            print.PrintExcel();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection rowColl;
            rowColl = dataGridView1.SelectedRows;
            if (rowColl.Count == 0)
            {
                MessageBox.Show("Не выбрано ни одной строки для копирования!");
                return;
            }
            string select;
            var settings = Properties.Settings.Default;
            string birth = "";
            int idSex = 5;
            string indx = "";
            int id_city = 62;
            string country = "";
            int id_district = 17;
            string street = "";
            string house = "";
            string flat = "";
            string phone = "";
            string passpSeries = "";
            string passpNumber = "";
            string passpKemVidan = "";
            string passpDate = "";
            int idGroup = 935;
            int id_otdelenie = 5;
            int id_comm = 3;
            string prikazNumIn = "";
            string dateIn = "";
            string prikazNumOut = "";
            string dateOut = "";
            string prichinaOut = "";
            string kval = "";
            string prikazNumKval = "";
            string note = "";
            int id_goden = 3;
            string katGodnost = "";
            int id_academ = 2;
            string ciizenship = "";
            string homePhone = "";
            try
            {
                foreach (DataGridViewRow row in rowColl)
                {
                    string name = row.Cells[0].Value.ToString();
                    if (settings.birth == true)
                        birth = row.Cells[1].Value.ToString();
                    idSex = ConvertSexValueToIdSex(row.Cells[2].Value.ToString());
                    if (settings.indx == true)
                        indx = row.Cells[5].Value.ToString();
                    if (settings.city == true)
                        id_city = ConvertCityNameToIdCity(row.Cells[6].Value.ToString());
                    if (settings.country == true)
                        country = row.Cells[8].Value.ToString();
                    if (settings.district == true)
                        id_district = ConvertDistrictNameToIdDistrict(row.Cells[7].Value.ToString());
                    if (settings.street == true)
                        street = row.Cells[9].Value.ToString();
                    if (settings.house == true)
                        house = row.Cells[10].Value.ToString();
                    if (settings.flat == true)
                        flat = row.Cells[11].Value.ToString();
                    if (settings.phone == true)
                        phone = row.Cells[12].Value.ToString();
                    if (settings.passpSeries == true)
                        passpSeries = row.Cells[15].Value.ToString();
                    if (settings.passpNumber == true)
                        passpNumber = row.Cells[16].Value.ToString();
                    if (settings.passpKemVidan == true)
                        passpKemVidan = row.Cells[17].Value.ToString();
                    if (settings.passpDate == true)
                        passpDate = row.Cells[18].Value.ToString();
                    if (settings.groupName == true)
                        idGroup = ConvertGroupNameToIdGroup(row.Cells[19].Value.ToString());
                    if (settings.otdelenie == true)
                        id_otdelenie = ConvertOtdelenieNumberToIdOtdelenie(row.Cells[20].Value.ToString());
                    if (settings.comm == true)
                        id_comm = ConvertCommTypeToIdComm(row.Cells[21].Value.ToString());
                    if (settings.prikazNumIn == true)
                        prikazNumIn = row.Cells[22].Value.ToString();
                    if (settings.dateIn == true)
                        dateIn = row.Cells[23].Value.ToString();
                    if (settings.prikazNumOut == true)
                        prikazNumOut = row.Cells[24].Value.ToString();
                    if (settings.dateOut == true)
                        dateOut = row.Cells[25].Value.ToString();
                    if (settings.prichinaOut == true)
                        prichinaOut = row.Cells[26].Value.ToString();
                    if (settings.kval == true)
                        kval = row.Cells[27].Value.ToString();
                    if (settings.prikazNumKval == true)
                        prikazNumKval = row.Cells[28].Value.ToString();
                    if (settings.Note == true)
                        note = row.Cells[30].Value.ToString();
                    if (settings.goden == true)
                        id_goden = ConvertGodenToIdGoden(row.Cells[3].Value.ToString());
                    if (settings.katgodn == true)
                        katGodnost = row.Cells[4].Value.ToString();
                    if (settings.akademotpusk == true)
                        id_academ = ConvertAcademValueToIdAcadem(row.Cells[29].Value.ToString());
                    if (settings.citizenship == true)
                        ciizenship = row.Cells[14].Value.ToString();
                    if (settings.homePhone == true)
                        homePhone = row.Cells[13].Value.ToString();
                    select = @"INSERT INTO Student
                      (name, birth, id_sex, indx, id_city, country, id_district, street, house, flat, phone, passpSeries, passpNumber, 
                       passpKemVidan, passpDate, idGroup, id_otdelenie, id_comm, prikazNumIn, dateIn, prikazNumOut, dateOut, prichinaOut, 
                       tabNum, kval, prikazNumKval, kodDoc, note, id_goden, katGodnost, id_academ, ciizenship, homePhone)
                      VALUES ('" + name + "','" + birth + "'," + idSex + ",'" + indx + "'," + id_city + ",'" + country + "'," + id_district + @",
                              '" + street + "','" + house + "','" + flat + "','" + phone + "','" + passpSeries + "','" + passpNumber + "','" + passpKemVidan + @"',
                              '" + passpDate + "'," + idGroup + "," + id_otdelenie + "," + id_comm + ",'" + prikazNumIn + "','" + dateIn + "','" + prikazNumOut + @"',
                              '" + dateOut + "','" + prichinaOut + "','','" + kval + "','" + prikazNumKval + "','','" + note + "'," + id_goden + @",
                              '" + katGodnost + "'," + id_academ + ",'" + ciizenship + "','" + homePhone + "')";
                    Util.FillTable("Student", select);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("При копировании что-то пошло не так! \n" + ex.ToString());
            }
            
            LoadTable2();
        }

        private void удалитьВыделенныеСтрокиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection rowColl = dataGridView1.SelectedRows;
            if (rowColl.Count == 0)
            {
                MessageBox.Show("Не выбрано ни одной строки для удаления!");
                return;
            }
            string deleted;
            try
            {
                if (MessageBox.Show("Количество строк для удаления: " + rowColl.Count + ".\n Эти строки будут удалены. Вы уверены?", "Удаление строк", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    if (rowColl.Count == 1)
                        deleted = @"DELETE FROM Student WHERE id = '" + rowColl[0].Cells[31].Value + "'";
                    else
                    {
                        deleted = @"DELETE FROM Student WHERE id = '";
                        foreach (DataGridViewRow row in rowColl)
                        {
                            deleted += row.Cells[31].Value + "' OR id = '";

                        }
                        deleted += @"'";
                    }
                    Util.FillTable("Student_Delete", deleted);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("При удалении что-то пошло не так! \n" + ex.ToString());
            }

            LoadTable2();
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Setting setting = new Setting();
            setting.ShowDialog();
        }

        private void дляДипломовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            backgroundWorker3.RunWorkerAsync();
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            PrintForDiplom printDiplom = new PrintForDiplom();
            printDiplom.print();
        }
    }
}
