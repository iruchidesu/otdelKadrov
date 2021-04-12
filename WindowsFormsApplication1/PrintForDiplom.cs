using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Collections;
using System.IO;

namespace WindowsFormsApplication1
{
    class PrintForDiplom
    {

        public PrintForDiplom()
        {
        }

        ArrayList groups = new ArrayList();

        void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                GC.Collect();
            }
        }

        public ArrayList group = new ArrayList();

        public void print()
        {
            FileStream stream = File.Open("list.txt", FileMode.Open, FileAccess.Read);
            if (stream != null)
            {
                StreamReader reader = new StreamReader(stream);
                bool j = true;
                group = new ArrayList();
                while (j == true)
                {
                    group.Add(reader.ReadLine());
                    if (reader.EndOfStream)
                    {
                        j = false;
                    }
                }
                stream.Close();
            }

            foreach (string gr in group)
            {
                string select = @"SELECT Student.name
                              FROM Student INNER JOIN
                                   [Group] ON Student.idGroup = [Group].id
                              WHERE ([Group].groupName = '" + gr + @"' AND Student.prikazNumOut = '')";
                select += @" ORDER BY Student.name";

                DataSet ds1 = new DataSet();

                try
                {
                    ds1 = Util.FillTable("Students", select);
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString());
                }

                Object exMiss = System.Reflection.Missing.Value;
                Excel.Workbook exclBook;
                Excel.Worksheet exclSheet;
                Excel.Application exclApp = new Excel.Application();

                exclBook = exclApp.Workbooks.Add();
                exclSheet = (Excel.Worksheet)exclBook.Sheets[1];

                //exclApp.Visible = true;

                exclSheet.Cells[1, 1] = "Студенты";
                exclSheet.Cells[5, 1] = "№ п.п";
                exclSheet.Cells[5, 2] = "Фамилия";
                exclSheet.Cells[5, 3] = "Имя";
                exclSheet.Cells[5, 4] = "Отчество";
                exclSheet.Cells[5, 5] = "Рег. № диплома";
                exclSheet.Cells[5, 6] = "Рег. № сертификата";
                exclSheet.Cells[5, 7] = "Рег № диплома ПК";
                exclSheet.Cells[5, 8] = "Рег. № удостоверения";

                int rowCount = 5;
                int rowNumber = 0;
                string[] fio;

                foreach (DataRow str in ds1.Tables[0].Rows)
                {
                    rowCount++;
                    rowNumber++;
                    fio = new string[4];
                    fio = str.ItemArray[0].ToString().Split(' ');
                    exclSheet.Cells[rowCount, 1] = rowNumber.ToString();
                    exclSheet.Cells[rowCount, 2] = fio[0];
                    exclSheet.Cells[rowCount, 3] = fio[1];
                    try
                    {
                        exclSheet.Cells[rowCount, 4] = fio[2] + "-" + fio[3];
                    }
                    catch
                    {
                        exclSheet.Cells[rowCount, 4] = fio[2];
                    }
                }

                //exclApp.Visible = true;

                string path = @"D:\D\Группы\" + gr + @".xlsx";

                exclBook.SaveAs(path);
                exclBook.Close();
                exclApp.Quit();

                ReleaseObject(exclApp);
            }

            MessageBox.Show("Операция успешно завершена.");
        }
    }
}
