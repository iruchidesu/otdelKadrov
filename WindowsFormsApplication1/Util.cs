using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    class Util
    {
        public static DataSet FillTable(string dtable, string query)
        {
            DataSet dataset = new DataSet();
            try
            {
                SqlConnection connectionstring = new SqlConnection(Properties.Resources.connectionstring);
                SqlDataAdapter da = new SqlDataAdapter(query, connectionstring);
                da.Fill(dataset, dtable);
                return dataset;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
                return dataset;
            }
        }

        public static int CalcKurs(string groupName) //вычисление курса
        {
            int kursNumber;
            string str;
            int dateNowYear = DateTime.Now.Year;
            int dateNowMonth = DateTime.Now.Month;
            string[] arrayKurs = groupName.Split('-');
            try
            {
                str = "20" + arrayKurs[2];
            }
            catch
            {
                str = "20" + arrayKurs[1];
            }
            if (dateNowMonth <= 6)
                kursNumber = dateNowYear - int.Parse(str);
            else
                kursNumber = dateNowYear - int.Parse(str) + 1;

            if (groupName.Substring((groupName.Length - 5), 2) == "11")
                kursNumber += 1;
            if (kursNumber > 4)
                kursNumber = 4;
            if (groupName.StartsWith("ПУ"))
                kursNumber = 5;
            return kursNumber;
        }
    }
}
