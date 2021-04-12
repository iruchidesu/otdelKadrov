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
    public partial class PrintForm9 : Form
    {
        public PrintForm9()
        {
            InitializeComponent();
        }

        private void LoadDistrict() // загрузка районов в комбобокс
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

        private void LoadGroups() // загрузка групп в комбобокс
        {
            DataSet ds = new DataSet();
            string select = "SELECT groupName FROM [Group] ";
            try
            {
                ds = Util.FillTable("Group", select);
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

        private void printForm9_Load(object sender, EventArgs e)
        {
            LoadGroups();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void districtRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (districtRadioButton.Checked == true)
            {
                LoadDistrict();
                medCheckBox.Enabled = false;
                medCheckBox.Checked = false;
             }
            else
            {
                LoadGroups();
                medCheckBox.Enabled = true;
            }


        }

        private void printButton_Click(object sender, EventArgs e)
        {


            if (String.IsNullOrWhiteSpace(comboBox1.Text))//вставить проверку, что указаны годы
            {
                MessageBox.Show("Выберите группу или район", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


            if (medCheckBox.Checked == true)
            {
                pictureBox1.Visible = true;
                groupRadioButton.Enabled = false;
                districtRadioButton.Enabled = false;
                comboBox1.Enabled = false;
                medCheckBox.Enabled = false;
                printButton.Enabled = false;
                closeButton.Enabled = false;
                backgroundWorker2.RunWorkerAsync(Tuple.Create(groupRadioButton.Checked, districtRadioButton.Checked, comboBox1.Text, medCheckBox.Checked));
            }
            else
            {
                pictureBox1.Visible = true;
                groupRadioButton.Enabled = false;
                districtRadioButton.Enabled = false;
                comboBox1.Enabled = false;
                medCheckBox.Enabled = false;
                printButton.Enabled = false;
                closeButton.Enabled = false;
                backgroundWorker1.RunWorkerAsync(Tuple.Create(groupRadioButton.Checked, districtRadioButton.Checked, comboBox1.Text, medCheckBox.Checked));
            }


        }

        private int ConvertDistrictNameToIdDistrict(string strdistrict)
        {
            string select = @"SELECT id FROM district WHERE ( district = '" + strdistrict + "')";
            DataSet ds = Util.FillTable("iddistrict", select);
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        private int ConvertGroupNameToIdGroup(string groupname)
        {
            string select = @"SELECT id FROM [Group]
                              WHERE groupName = '" + groupname + "'";
            DataSet ds = Util.FillTable("idGroup", select);
            return (int)ds.Tables[0].Rows[0].ItemArray[0];
        }

        private void groupRadioButton_CheckedChanged(object sender, EventArgs e)
        {

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
            var t = e.Argument as Tuple<bool, bool, string, bool>;
            if (t.Item3 == "")
            {
                MessageBox.Show("Район не задан!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string select = "";
            if (t.Item2 == true)
            {

                select = @"SELECT     Student.name, Student.birth, city.city, district.district, Student.country, Student.street, Student.house, Student.flat, Student.phone
                                FROM         Student INNER JOIN
                                city ON Student.id_city = city.id INNER JOIN
                                district ON Student.id_district = district.id
                                WHERE id_district = " + ConvertDistrictNameToIdDistrict(t.Item3) +
                                @" AND (prikazNumKval = '') AND (prikazNumOut = '') 
                                ORDER BY name";
            }
            else
            {

                select = @"SELECT     Student.name, Student.birth, city.city, district.district, Student.country, Student.street, Student.house, Student.flat, Student.phone
                                                FROM         Student INNER JOIN
                                                city ON Student.id_city = city.id INNER JOIN
                                                district ON Student.id_district = district.id INNER JOIN
                                                [Group] ON Student.idGroup = [Group].id 
                                                WHERE idGroup = " + ConvertGroupNameToIdGroup(t.Item3) +
                                                @" AND (prikazNumKval = '') AND (prikazNumOut = '') 
                                                ORDER BY name";
            }

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

            int columnsCount = 6;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, ds1.Tables[0].Rows.Count + 3, columnsCount);
            tb.Columns[1].Width = 30;
            tb.Rows[1].Height = 40;
            tb.Columns[2].Width = 180;
            tb.Rows[2].Height = 40;
            tb.Columns[3].Width = 60;
            tb.Rows[3].Height = 40;
            tb.Columns[4].Width = 120;
            tb.Columns[5].Width = 70;

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
            if (districtRadioButton.Checked == true)
            {
                tb.Cell(2, 1).Range.Text = " студентов проживающих в " + DistrictForm(t.Item3) + " районе";
            }
            else
            {
                tb.Cell(2, 1).Range.Text = " студентов группы" + t.Item3;
            }

            tb.Rows[3].Select();
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(3, 1).Range.Text = "№ п/п";
            tb.Cell(3, 2).Range.Text = "ФИО";
            tb.Cell(3, 3).Range.Text = "Дата рождения";
            tb.Cell(3, 4).Range.Text = "Домашний адрес";
            tb.Cell(3, 5).Range.Text = "Телефон";
            tb.Cell(3, 6).Range.Text = " ";


            int rowCount = 3;
            int rowNumber = 0;

            foreach (DataRow str in ds1.Tables[0].Rows)
            {
                rowCount++;
                rowNumber++;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.ItemArray[0].ToString();
                tb.Cell(rowCount, 3).Range.Text = DateTime.Parse(str.ItemArray[1].ToString()).ToShortDateString();
                tb.Cell(rowCount, 4).Range.Text = str.ItemArray[2].ToString() + " " + str.ItemArray[3].ToString() + " р-н" + " ул. " + str.ItemArray[5].ToString() + " д. " +
                                                  str.ItemArray[6].ToString() + " кв. " + str.ItemArray[7].ToString();
                tb.Cell(rowCount, 5).Range.Text = str.ItemArray[8].ToString();
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
            wdApp.Selection.InsertBreak(6);

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

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            var t = e.Argument as Tuple<bool, bool, string, bool>;
            if (t.Item3 == "")
            {
                MessageBox.Show("Группа не задана!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //здесь формирование ворд-документа для медчасти
            string select = @"SELECT     Student.name, Student.birth, city.city, district.district, Student.country, Student.street, Student.house, Student.flat, Student.phone
                                                FROM         Student INNER JOIN
                                                city ON Student.id_city = city.id INNER JOIN
                                                district ON Student.id_district = district.id INNER JOIN
                                                [Group] ON Student.idGroup = [Group].id 
                                                WHERE idGroup = " + ConvertGroupNameToIdGroup(t.Item3) +
                                                @" AND (prikazNumKval = '') AND (prikazNumOut = '') 
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
            wdDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            wdDoc.PageSetup.LeftMargin = 40;
            wdDoc.PageSetup.RightMargin = 25;
            wdDoc.PageSetup.TopMargin = 20;
            wdDoc.PageSetup.BottomMargin = 20;
            Word.Table tb;
            Word.Range _range;

            int columnsCount = 7;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, ds1.Tables[0].Rows.Count + 3, columnsCount);
            tb.Columns[1].Width = 30;
            tb.Rows[1].Height = 25;
            tb.Columns[2].Width = 210;
            tb.Rows[2].Height = 40;
            tb.Columns[3].Width = 60;
            tb.Rows[3].Height = 40;
            tb.Columns[4].Width = 150;
            tb.Columns[5].Width = 60;
            tb.Columns[6].Width = 60;
            tb.Columns[7].Width = 210;

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
            tb.Cell(2, 1).Range.Text = " студентов группы " + t.Item3;

            tb.Rows[3].Select();
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(3, 1).Range.Text = "№ п/п";
            tb.Cell(3, 2).Range.Text = "ФИО";
            tb.Cell(3, 3).Range.Text = "Дата рождения";
            tb.Cell(3, 4).Range.Text = "Домашний адрес";
            tb.Cell(3, 5).Range.Text = "Группа здоров.";
            tb.Cell(3, 6).Range.Text = "Физк. группа";
            tb.Cell(3, 7).Range.Text = "Диагноз";


            int rowCount = 3;
            int rowNumber = 0;

            foreach (DataRow str in ds1.Tables[0].Rows)
            {
                rowCount++;
                rowNumber++;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.ItemArray[0].ToString();
                tb.Cell(rowCount, 3).Range.Text = DateTime.Parse(str.ItemArray[1].ToString()).ToShortDateString();
                tb.Cell(rowCount, 4).Range.Text = str.ItemArray[2].ToString() + "   " + str.ItemArray[3].ToString() + " р-н" + " ул. " + str.ItemArray[5].ToString() + " д. " +
                                                  str.ItemArray[6].ToString() + " кв. " + str.ItemArray[7].ToString();
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
            groupRadioButton.Enabled = true;
            districtRadioButton.Enabled = true;
            comboBox1.Enabled = true;
            medCheckBox.Enabled = true;
            printButton.Enabled = true;
            closeButton.Enabled = true;
            if (districtRadioButton.Checked == true)
                medCheckBox.Enabled = false;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
            groupRadioButton.Enabled = true;
            districtRadioButton.Enabled = true;
            comboBox1.Enabled = true;
            printButton.Enabled = true;
            closeButton.Enabled = true;
            if (districtRadioButton.Checked == true)
                medCheckBox.Enabled = false;
            
        }
    }
}
