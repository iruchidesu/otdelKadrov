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
    public partial class PrintGroup : Form
    {
        public PrintGroup()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void PrintGroup_Load(object sender, EventArgs e)
        {
            LoadGroups();
        }

        private void LoadGroups() // загрузка групп в комбобоксы
        {
            DataSet ds = new DataSet();
            string select = "SELECT groupName FROM [Group] ";
            try
            {
                ds = Util.FillTable("Student", select);
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
            comboBox1.SelectedItem = comboBox1.Items[1];
        }

        private int ConvertGroupNameToIdGroup(string groupname)
        {
            string select = @"SELECT id FROM [Group]
                              WHERE groupName = '" + groupname + "'";
            DataSet ds = Util.FillTable("idGroup", select);
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            noBirthCheckBox.Enabled = false;
            comboBox1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            backgroundWorker1.RunWorkerAsync(Tuple.Create(comboBox1.Text, noBirthCheckBox.Checked));
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
                button1.Enabled = false;
            else
                button1.Enabled = true;
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
                button1.Enabled = false;
            else
                button1.Enabled = true;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var t = e.Argument as Tuple<string, bool>;
            if (String.IsNullOrWhiteSpace(t.Item1)) //вставить проверку, что указаны годы
            {
                MessageBox.Show("Не указаны годы рождения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            string select = @"SELECT Student.name";
            if (t.Item2 == false)
                select += ", Student.birth";
            select += @" FROM Student INNER JOIN
                             [Group] ON Student.idGroup = [Group].id 
                        WHERE idGroup = " + ConvertGroupNameToIdGroup(t.Item1) + @" AND (prikazNumKval = '') AND (prikazNumOut = '') 
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
            wdDoc.PageSetup.TopMargin = 50;
            wdDoc.PageSetup.BottomMargin = 20;
            Word.Table tb;
            Word.Range _range;

            int columnsCount = 0;

            if (t.Item2 == true)
                columnsCount = 3;
            else
                columnsCount = 4;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, ds1.Tables[0].Rows.Count + 2, columnsCount);
            tb.Columns[1].Width = 40;
            tb.Rows[1].Height = 30;
            tb.Columns[2].Width = 300;
            tb.Rows[2].Height = 20;
            if (t.Item2 == true)
            {
                tb.Columns[3].Width = 186;
            }
            else
            {
                tb.Columns[3].Width = 86;
                tb.Columns[4].Width = 100;
            }

            Word.Row row = tb.Rows[1];
            Word.Cell firstCell = row.Cells[1];
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
            tb.Cell(1, 1).Range.Text = "Группа " + t.Item1;

            tb.Rows[2].Select();
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(2, 1).Range.Text = "№ п/п";
            tb.Cell(2, 2).Range.Text = "Ф.И.О.";
            if (t.Item2 == true)
            {
                tb.Cell(2, 3).Range.Text = "Примечание";
            }
            else
            {
                tb.Cell(2, 3).Range.Text = "Дата рожд.";
                tb.Cell(2, 4).Range.Text = "Примечание";
            }

            int rowCount = 2;
            int rowNumber = 0;

            foreach (DataRow str in ds1.Tables[0].Rows)
            {
                rowCount++;
                rowNumber++;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.ItemArray[0].ToString();
                if (t.Item2 == false)
                {
                    try
                    {
                        tb.Cell(rowCount, 3).Range.Text = DateTime.Parse(str.ItemArray[1].ToString()).ToShortDateString();
                    }
                    catch
                    {
                        tb.Cell(rowCount, 3).Range.Text = str.ItemArray[1].ToString();
                    }
                }
            }

            _range = wdDoc.Range(tb.Cell(3, 2).Range.Start, tb.Cell(rowCount, 2).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            _range = wdDoc.Range(tb.Cell(3, columnsCount).Range.Start, tb.Cell(rowCount, columnsCount).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;


            _range = wdDoc.Range(tb.Cell(2, 1).Range.Start, tb.Cell(rowCount, columnsCount).Range.End);
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
            noBirthCheckBox.Enabled = true;
            comboBox1.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
        }
    }
}