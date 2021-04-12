using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApplication1
{
    public partial class StringEditor : Form
    {
        public StringEditor()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void LoadTableCity()
        {
            string select = @"SELECT city, id FROM city";

            try
            {
                cityDataGrid.DataSource = Util.FillTable("city", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            cityDataGrid.DataMember = "city";
            cityDataGrid.Columns[0].HeaderText = "Город(Область)";
            cityDataGrid.Columns[1].Visible = false;
        }

        private void LoadTableDistricts()
        {
            string select = @"SELECT district, id FROM district";

            try
            {
                districtDataGrid.DataSource = Util.FillTable("district", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            districtDataGrid.DataMember = "district";
            districtDataGrid.Columns[0].HeaderText = "Район";
            districtDataGrid.Columns[1].Visible = false;
        }

        private void StringEditor_Load(object sender, EventArgs e)
        {
            LoadTableCity();
            LoadTableDistricts();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection c;
            c = cityDataGrid.SelectedRows;
            if (c.Count == 0)
            {
                MessageBox.Show("Не выделено ни одной строки");
            }
            else
            {
                try
                {
                    Util.FillTable("city", "DELETE FROM city WHERE id = '" + c[0].Cells[1].Value + "'");
                }
                catch
                {
                    MessageBox.Show("Невозможно удалить, т.к. это значение используется!");
                }
            }
            LoadTableCity();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection c;
            c = districtDataGrid.SelectedRows;
            if (c.Count == 0)
            {
                MessageBox.Show("Не выделено ни одной строки");
            }
            else
            {
                try
                {
                    Util.FillTable("district", "DELETE FROM district WHERE id = '" + c[0].Cells[1].Value + "'");
                }
                catch
                {
                    MessageBox.Show("Невозможно удалить, т.к. это значение используется!");
                }
            }
            LoadTableDistricts();
        }
    }
}
