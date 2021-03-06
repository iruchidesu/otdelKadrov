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
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


namespace WindowsFormsApplication1
{
    public partial class PrintDistrict : Form
    {
        public PrintDistrict()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void PrintDistrict_Load(object sender, EventArgs e)
        {
            LoadDistricts();
        }

        private int ConvertDistrictNameToIdDistrict(string strdistrict)
        {
            string select = @"SELECT id FROM district WHERE ( district = '" + strdistrict + "')";
            DataSet ds = Util.FillTable("iddistrict", select);
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            textBox1.Enabled = false;
            comboBox1.Enabled = false;
            textBox2.Enabled = false;
            dateTimePicker1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            backgroundWorker1.RunWorkerAsync(Tuple.Create(comboBox1.Text, textBox1.Text, textBox2.Text, dateTimePicker1.Text));
        }

        private void LoadDistricts() // загрузка районов в комбобокс
        {
            DataSet ds = new DataSet();
            string select = "SELECT district FROM district ";
            try
            {
                ds = Util.FillTable("district", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            comboBox1.Items.Clear();
            foreach (DataRow itm in ds.Tables[0].Rows)
            {
                comboBox1.Items.Add(itm.ItemArray[0]);
            }
            comboBox1.SelectedItem = comboBox1.Items[0];
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
                button2.Enabled = false;
            else
                button2.Enabled = true;
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
                button2.Enabled = false;
            else
                button2.Enabled = true;
        }

        private string DistrictForm(string distric)
        {
            string changedName = "";
            string tempDistrict = "";
            tempDistrict = distric.Substring(0, distric.Length - 2);
            changedName = tempDistrict + "ом";

            return changedName;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var t = e.Argument as Tuple<string, string, string, string>;
            if (String.IsNullOrWhiteSpace(t.Item2) || String.IsNullOrWhiteSpace(t.Item3)) //вставить проверку, что указаны годы
            {
                MessageBox.Show("Не указаны годы рождения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (t.Item1 == "")
            {
                MessageBox.Show("Район не задан!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string select = @"SELECT     Student.name, Student.birth, city.city, Student.street, Student.house, Student.flat, district.district
                                FROM         Student INNER JOIN
                                city ON Student.id_city = city.id INNER JOIN
                                district ON Student.id_district = district.id
                                WHERE id_district = " + ConvertDistrictNameToIdDistrict(t.Item1) +
                                @" AND (prikazNumKval = '') AND (prikazNumOut = '') 
                                   AND (birth >= '01.01." + t.Item2 + "' AND birth <= '31.12." + t.Item3 + @"')
                                ORDER BY name";

            DataSet ds1 = new DataSet();

            try
            {
                ds1 = Util.FillTable("Student", select);
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (ds1.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show("С указанными данными никого нет!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Word.Application wdApp = new Word.Application();
            Word.Document wdDoc = new Word.Document();
            Object wdMiss = System.Reflection.Missing.Value;

            wdDoc = wdApp.Documents.Add(ref wdMiss, ref wdMiss, ref wdMiss, ref wdMiss);
            //wdApp.Visible = true; //сначала формируется документ, показывать потом
            wdDoc.PageSetup.LeftMargin = 40;
            wdDoc.PageSetup.RightMargin = 25;
            wdDoc.PageSetup.TopMargin = 20;
            wdDoc.PageSetup.BottomMargin = 20;
            Word.Table tb;
            Word.Range _range;

            int columnsCount = 5;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, ds1.Tables[0].Rows.Count + 3, columnsCount);
            tb.Columns[1].Width = 40;
            tb.Rows[1].Height = 40;
            tb.Columns[2].Width = 210;
            tb.Rows[2].Height = 40;
            tb.Columns[3].Width = 60;
            tb.Rows[3].Height = 40;
            tb.Columns[4].Width = 140;
            tb.Columns[5].Width = 90;

            Word.Row row = tb.Rows[1];
            Word.Cell firstCell = row.Cells[1];
            foreach (Word.Cell currCell in row.Cells)
            {
                if (currCell.ColumnIndex != firstCell.ColumnIndex)
                {
                    firstCell.Merge(currCell);
                }
            }
            row = tb.Rows[2];
            firstCell = row.Cells[1];
            foreach (Word.Cell currCell in row.Cells)
            {
                if (currCell.ColumnIndex != firstCell.ColumnIndex)
                {
                    firstCell.Merge(currCell);
                }
            }

            tb.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wdApp.Selection.Range.Font.Name = "Times New Roman";
            wdApp.Selection.Range.Font.Size = 10;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 0;

            tb.Cell(1, 1).Select();
            wdApp.Selection.Range.Font.Size = 14;
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(1, 1).Range.Text = "Список";
            tb.Rows[2].Select();
            wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(2, 1).Range.Text = " студентов " + t.Item2 + " - " + t.Item3 +
                                       " г.р. проживающих в " + DistrictForm(t.Item1) + " районе на " + t.Item4 + " года ";

            tb.Rows[3].Select();
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(3, 1).Range.Text = "№ п/п";
            tb.Cell(3, 2).Range.Text = "ФИО";
            tb.Cell(3, 3).Range.Text = "Дата рождения";
            tb.Cell(3, 4).Range.Text = "Домашний адрес";
            tb.Cell(3, 5).Range.Text = "Примечание";


            int rowCount = 3;
            int rowNumber = 0;

            foreach (DataRow str in ds1.Tables[0].Rows)
            {
                rowCount++;
                rowNumber++;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.ItemArray[0].ToString();
                tb.Cell(rowCount, 3).Range.Text = DateTime.Parse(str.ItemArray[1].ToString()).ToShortDateString(); //номер курса считать
                tb.Cell(rowCount, 4).Range.Text = "г." + str.ItemArray[2].ToString() + ", ул." + str.ItemArray[3].ToString() + ", д." +
                                                  str.ItemArray[4].ToString() + ", кв." + str.ItemArray[5].ToString();
            }

            _range = wdDoc.Range(tb.Cell(4, 2).Range.Start, tb.Cell(rowCount, 5).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            _range = wdDoc.Range(tb.Cell(3, 1).Range.Start, tb.Cell(rowCount, columnsCount).Range.End);
            _range.Select();

            /* вызов макроса для отображения границ таблицы
            Sub Сетка()
            'Сетка макрос
            With Selection.Borders(wdBorderTop)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
            End With
            With Selection.Borders(wdBorderLeft)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
            End With
            With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
            End With
            With Selection.Borders(wdBorderRight)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
            End With
            With Selection.Borders(wdBorderHorizontal)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
            End With
            With Selection.Borders(wdBorderVertical)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
            End With
            End Sub 
            */
            try
            {
                wdApp.Run("Сетка");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " Сетка", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


            //таблица для подписей
            wdApp.Selection.EndOf(Word.WdUnits.wdStory);
            wdApp.Selection.InsertBreak(10);

            Word.Table tb2;
            tb2 = wdDoc.Tables.Add(wdApp.Selection.Range, 2, 4);

            tb2.Columns[1].Width = 40;
            tb2.Columns[2].Width = 180;
            tb2.Columns[3].Width = 180;

            tb2.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wdApp.Selection.Range.Font.Name = "Times New Roman";
            wdApp.Selection.Range.Font.Size = 10;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 0;

            tb2.Rows[1].Select();
            tb2.Cell(1, 2).Range.Text = "Директор колледжа";
            tb2.Cell(1, 4).Range.Text = " ";

            tb2.Rows[2].Select();
            tb2.Cell(2, 2).Range.Text = " ";
            tb2.Cell(2, 3).Range.Text = " ";

            //нумерация страниц
            Word.Window activeWindow = wdDoc.Application.ActiveWindow;
            object currentPage = Word.WdFieldType.wdFieldPage;
            object totalPages = Word.WdFieldType.wdFieldNumPages;
            //переход к редактированию футера
            activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            activeWindow.ActivePane.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            //напечатает номер страницы в формате X стр. из Y
            activeWindow.Selection.Fields.Add(activeWindow.Selection.Range, ref currentPage, ref wdMiss, ref wdMiss);
            activeWindow.Selection.TypeText(" стр. из ");
            activeWindow.Selection.Fields.Add(activeWindow.Selection.Range, ref totalPages, ref wdMiss, ref wdMiss);
            //выход из футера
            activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            wdApp.Visible = true; //показать документ пользователю
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
            textBox1.Enabled = true;
            comboBox1.Enabled = true;
            textBox2.Enabled = true;
            dateTimePicker1.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
        }
    }

}