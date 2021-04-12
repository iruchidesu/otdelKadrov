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
    public partial class PrintForm10 : Form
    {
        public PrintForm10()
        {
            InitializeComponent();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            Hide();
        }

        private string Voenkomat(string district)
        {
            string str = "";
            switch (district)
            {
                case "Мошковский":
                    str = "Мошковский";
                    break;
                case "Купинский":
                    str = "Купинский";
                    break;
                case "Искитимский":
                    str = "Искитимский";
                    break;
                case "Бердский":
                    str = "Бердский";
                    break;
                case "Ордынский":
                    str = "Ордынский";
                    break;
                case "Сузунский":
                    str = "Сузунский";
                    break;
                case "Чулымский":
                    str = "Чулымский";
                    break;
                case "Карасукский":
                case "Баганский":
                    str = "Карасукский и Баганский";
                    break;
                case "Каргатский":
                case "Убинский":
                    str = "Каргатский и Убинский";
                    break;
                case "Коченевский":
                case "Колыванский":
                    str = "Коченевский и Колыванский";
                    break;
                case "Краснозерский":
                case "Доволенский":
                case "Кочковский":
                    str = "Краснозерский, Доволенский и Кочковский";
                    break;
                case "Куйбышевский":
                case "Северный":
                    str = "Куйбышевский и Северный";
                    break;
                case "Татарский":
                case "Усть-Таркский":
                case "Чистоозерный":
                    str = "Татарский, Усть-Таркский и Чистоозерный";
                    break;
                case "Тогучинский":
                case "Болотнинский":
                    str = "Тогучинский и Болотнинский";
                    break;
                case "Чановский":
                case "Венгеровский":
                case "Кыштовский":
                    str = "Чановский, Венгеровский и Кыштовский";
                    break;
                case "Черепановский":
                case "Маслянинский":
                    str = "Черепановский и Маслянинский";
                    break;
                case "Барабинский":
                case "Здвинский":
                    str = "Барабинский и Здвинский";
                    break;
                case "Новосибирский":
                    str = "Новосибирский р-н, г. Обь и р.п. Кольцово";
                    break;
                case "Советский":
                case "Первомайский":
                    str = "Советский и Первомайский";
                    break;
                case "Заельцовский":
                case "Центральный":
                case "Железнодорожный":
                case "Октябрьский":
                    str = "Октябрьский р-он и центральный административный округ";
                    break;
                case "Кировский":
                case "Ленинский":
                    str = "Кировский и Ленинский";
                    break;
                case "Дзержинский":
                case "Калининский":
                    str = "Дзержинский и Калининский";
                    break;
            }
            return str;
        }

        private void printButton_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            birthTextBox1.Enabled = false;
            birthTextBox2.Enabled = false;
            dateTimePicker1.Enabled = false;
            printButton.Enabled = false;
            closeButton.Enabled = false;
            backgroundWorker1.RunWorkerAsync(Tuple.Create(birthTextBox1.Text, birthTextBox2.Text, dateTimePicker1.Text));
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var t = e.Argument as Tuple<string, string, string>;
            if (String.IsNullOrWhiteSpace(t.Item1) || String.IsNullOrWhiteSpace(t.Item2)) //вставить проверку, что указаны годы
            {
                MessageBox.Show("Не указаны годы рождения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //здесь формирование ворд-документа
            string select = @"SELECT Student.name, Student.birth, district.district, Student.dateOut, sex.sex, Student.prichinaOut 
                              FROM Student INNER JOIN
                                   district ON Student.id_district = district.id INNER JOIN
                                   sex ON Student.id_sex = sex.id
                              WHERE ((prikazNumKval = '') AND (prikazNumOut != '')) AND
                                    (sex.sex = 'муж.') AND (Student.dateOut >= '" + t.Item3 + @"') AND (Student.id_goden = 1) AND 
                                    ((Student.id_city = 63) OR (Student.id_city = 64)) AND ((Student.birth >= '01.01." + t.Item1 + @"') AND (Student.birth <= '31.12." + birthTextBox2.Text + @"'))
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

            int columnsCount = 6;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, ds1.Tables[0].Rows.Count + 3, columnsCount);
            tb.Columns[1].Width = 30;
            tb.Columns[2].Width = 200;
            tb.Columns[3].Width = 40;
            tb.Columns[4].Width = 140;
            tb.Columns[5].Width = 70;
            tb.Columns[6].Width = 70;

            tb.Rows[2].Height = 30;

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
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wdApp.Selection.Range.Font.Name = "Times New Roman";
            wdApp.Selection.Range.Font.Size = 10;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 0;

            tb.Cell(1, 1).Select();
            wdApp.Selection.Range.Font.Size = 12;
            tb.Cell(1, 1).Range.Text = "Список граждан, подлежащих призыву на военную службу и отчисленных из образовательных учреждений среднего профессионального образования по военному комиссариату Дзержинского и Калининского районов г.Новосибирска с " + t.Item3 + " года. \n";
            tb.Rows[2].Select();
            wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tb.Cell(2, 1).Range.Text = " ";

            tb.Rows[3].Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(3, 1).Range.Text = "№ п/п";
            tb.Cell(3, 2).Range.Text = "Фамилия, имя, отчество";
            tb.Cell(3, 3).Range.Text = "Год рожд.";
            tb.Cell(3, 4).Range.Text = "Военный комиссариат";
            tb.Cell(3, 5).Range.Text = "Дата отчисления";
            tb.Cell(3, 6).Range.Text = "Примечание";

            int rowCount = 3;
            int rowNumber = 0;

            foreach (DataRow str in ds1.Tables[0].Rows)
            {
                rowCount++;
                rowNumber++;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.ItemArray[0].ToString();
                tb.Cell(rowCount, 3).Range.Text = DateTime.Parse(str.ItemArray[1].ToString()).Year.ToString();
                tb.Cell(rowCount, 4).Range.Text = Voenkomat(str.ItemArray[2].ToString()); //вместо района пишется соответсвующий военкомат, если район не соответствует определенным то будет пусто
                tb.Cell(rowCount, 5).Range.Text = DateTime.Parse(str.ItemArray[3].ToString()).ToShortDateString();
                tb.Cell(rowCount, 6).Range.Text = str.ItemArray[5].ToString();
            }

            _range = wdDoc.Range(tb.Cell(4, 5).Range.Start, tb.Cell(rowCount, 5).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
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
            wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 0;

            tb2.Rows[1].Select();
            tb2.Cell(1, 2).Range.Text = "Директор колледжа \n";
            tb2.Cell(1, 4).Range.Text = "  \n";

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
            birthTextBox1.Enabled = true;
            birthTextBox2.Enabled = true;
            dateTimePicker1.Enabled = true;
            printButton.Enabled = true;
            closeButton.Enabled = true;
        }
    }
}
