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
    public partial class PrintForm1 : Form
    {
        public PrintForm1()
        {
            InitializeComponent();
        }

        string source = Properties.Resources.connectionstring;

        private DataSet fillTable(string dtable, string query)
        {
            DataSet ds1 = new DataSet();
            try
            {
                SqlConnection connectionstring = new SqlConnection(source);
                SqlDataAdapter da = new SqlDataAdapter(query, connectionstring);
                da.Fill(ds1, dtable);
                return ds1;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return ds1;
            }
        }

        private void PrintForm1_Load(object sender, EventArgs e)
        {
            load_grp(); // загрузка групп в комбобоксы
        }

        private void load_grp() // загрузка групп в комбобоксы
        {
            DataSet ds1 = new DataSet();
            string select = "SELECT nameGroup FROM [Group] ";
            try
            {
                ds1 = fillTable("Group", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            comboBox1.Items.Clear();
            foreach (DataRow itm in ds1.Tables[0].Rows)
            {
                comboBox1.Items.Add(itm.ItemArray[0]);
            }
        }
    }
}
