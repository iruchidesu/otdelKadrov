using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;

namespace WindowsFormsApplication1
{
    public partial class Group : Form
    {
        public Group()
        {
            InitializeComponent();
        }

        private void LoadTable()
        {
            string select = @"SELECT groupName, id FROM [Group] ";
            DataSet ds = Util.FillTable("Group", select);

            dataGridView2.AutoGenerateColumns = false;
            dataGridView2.DataMember = "Group";
            dataGridView2.DataSource = ds;

            CalcKurs(ds); //вычисление курса

        }

        public void CalcKurs(DataSet dataSet) //вычисление курса
        {
            int dateNowYear = DateTime.Now.Year;
            int dateNowMonth = DateTime.Now.Month;
            ArrayList arraylistKurs = new ArrayList();
            string[] arrayKurs = new string[3];
            int count = 0;
            int k;
            foreach (DataRow row in dataSet.Tables[0].Rows)
            {
                if (row.ItemArray[0].ToString() == "")
                {
                    continue;
                }
                else
                {
                    arrayKurs = row.ItemArray[0].ToString().Split('-');
                    try
                    {
                        arraylistKurs.Add("20" + arrayKurs[2]);
                    }
                    catch
                    {
                        arraylistKurs.Add("20" + arrayKurs[1]);
                    }
                }
            }
            foreach (string str in arraylistKurs)
            {
                if (dateNowMonth <= 6)
                    k = dateNowYear - int.Parse(str);
                else
                    k = dateNowYear - int.Parse(str) + 1;
                if (dataGridView2.Rows[count].Cells[0].Value.ToString() != "")
                {
                    if (dataGridView2.Rows[count].Cells[0].Value.ToString().Substring((dataGridView2.Rows[count].Cells[0].Value.ToString().Length - 5), 2) == "11")
                        k += 1;
                }
                else
                    count++;
                if (k > 4)
                    k = 4;
                if (dataGridView2.Rows[count].Cells[0].Value.ToString().StartsWith("ПУ"))
                    k = 5;
                dataGridView2.Rows[count].Cells[1].Value = k.ToString();
                count++;
            }
        }

        private void Group_Load(object sender, EventArgs e)
        {
            LoadTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection c;
            c = dataGridView2.SelectedRows;
            if (c.Count == 0)
            {
                MessageBox.Show("Не выделено ни одной строки");
            }
            else
            {
                try
                {
                    Util.FillTable("Group", "DELETE FROM [Group] WHERE id = '" + c[0].Cells[2].Value + "'");
                }
                catch
                {
                    MessageBox.Show("Невозможно удалить, т.к. это значение используется!");
                }
            }
            LoadTable();
        }

        private void dataGridView2_Sorted(object sender, EventArgs e)
        {
            string select;
            if (dataGridView2.SortOrder == System.Windows.Forms.SortOrder.Ascending)
                select = @"SELECT groupName, id FROM [Group] ORDER BY groupName";
            else
                select = @"SELECT groupName, id FROM [Group] ORDER BY groupName DESC";
            DataSet ds = Util.FillTable("Group", select);
            CalcKurs(ds);
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentRow.Cells[2].Value.ToString() == "")
            {
                //int count = dataGridView1.Rows.Count - 1;
                string select = "INSERT INTO [Group] (groupName) VALUES ('" + dataGridView2.CurrentRow.Cells[0].Value + "')";
                Util.FillTable("[Group]", select);
            }
            else
            {
                string select = "UPDATE [Group] SET";
                switch (e.ColumnIndex)
                {
                    case 0:
                        select += " groupName ";
                        break;
                }
                select += " = '" + dataGridView2.CurrentCell.Value + "' WHERE (id = '" + dataGridView2.CurrentRow.Cells[2].Value + "')";

                try
                {
                    Util.FillTable("[Group]", select);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

   }
}
    


