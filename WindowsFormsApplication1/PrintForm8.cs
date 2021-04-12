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
    public partial class PrintForm8 : Form
    {
        public PrintForm8()
        {
            InitializeComponent();
        }
        string voenkomat = "";

        private void LoadVoenkomat() // загрузка военкоматов в комбобокс
        {
            string[] dist = new string[22];
            dist[0] = "Дзержинский и Калининский";
            dist[1] = "Кировский и Ленинский";
            dist[2] = "Октябрьский р-он и центральный административный округ";
            dist[3] = "Советский и Первомайский";
            dist[4] = "Новосибирский р-н, г. Обь и р.п. Кольцово";
            dist[5] = "Барабинский и Здвинский";
            dist[6] = "Бердский";
            dist[7] = "Искитимский";
            dist[8] = "Купинский";
            dist[9] = "Мошковский";
            dist[10] = "Ордынский";
            dist[11] = "Сузунский";
            dist[12] = "Чулымский";
            dist[13] = "Карасукский и Баганский";
            dist[14] = "Каргатский и Убинский";
            dist[15] = "Коченевский и Колыванский";
            dist[16] = "Краснозерский, Доволенский и Кочковский";
            dist[17] = "Куйбышевский и Северный";
            dist[18] = "Татарский, Усть-Таркский и Чистоозерный";
            dist[19] = "Тогучинский и Болотнинский";
            dist[20] = "Чановский, Венгеровский и Кыштовский";
            dist[21] = "Черепановский и Маслянинский";

            districtComboBox.Items.Clear();
            foreach (string itm in dist)
            {
                districtComboBox.Items.Add(itm);
            }
            districtComboBox.SelectedItem = districtComboBox.Items[0];
        }

        private string ConvertVoencomatNameToFullName(string voencomatName)
        {
            string str = "";
            switch (voencomatName)
            {
                case "Дзержинский и Калининский":
                    str = "Калининского и Дзержинского районов г.Новосибирска";
                    break;
                case "Кировский и Ленинский":
                    str = "Кировского и Ленинского районов г.Новосибирска";
                    break;
                case "Октябрьский р-он и центральный административный округ":
                    str = "Октябрьского района и центрального административного округа г.Новосибирска";
                    break;
                case "Советский и Первомайский":
                    str = "Советского и Первомайского районов г.Новосибирска";
                    break;
                case "Новосибирский р-н, г. Обь и р.п. Кольцово":
                    str = "Новосибирского района, г. Обь и р.п. Кольцово";
                    break;
                case "Барабинский и Здвинский":
                    str = "Барабинского и Здвинского районов";
                    break;
                case "Бердский":
                    str = "Бердского района";
                    break;
                case "Искитимский":
                    str = "Искитимского района";
                    break;
                case "Купинский":
                    str = "Купинского района";
                    break;
                case "Мошковский":
                    str = "Мошковского района";
                    break;
                case "Ордынский":
                    str = "Ордынского района";
                    break;
                case "Сузунский":
                    str = "Сузунского района";
                    break;
                case "Чулымский":
                    str = "Чулымского района";
                    break;
                case "Карасукский и Баганский":
                    str = "Карасукского и Баганского районов";
                    break;
                case "Каргатский и Убинский":
                    str = "Каргатского и Убинского районов";
                    break;
                case "Коченевский и Колыванский":
                    str = "Коченевского и Колыванского районов";
                    break;
                case "Краснозерский, Доволенский и Кочковский":
                    str = "Краснозерского, Доволенского и Кочковского районов";
                    break;
                case "Куйбышевский и Северный":
                    str = "Куйбышевского и Северного районов";
                    break;
                case "Татарский, Усть-Таркский и Чистоозерный":
                    str = "Татарского, Усть-Таркского и Чистоозерного районов";
                    break;
                case "Тогучинский и Болотнинский":
                    str = "Тогучинского и Болотнинского районов";
                    break;
                case "Чановский, Венгеровский и Кыштовский":
                    str = "Чановского, Венгеровского и Кыштовского районов";
                    break;
                case "Черепановский и Маслянинский":
                    str = "Черепановского и Маслянинского районов";
                    break;
            }
            return str;
        }

        private string ConvertDistrictNameToDistrictQuery(string districtName)
        {
            string sel = "";
            switch (districtName)
            {
                case "Дзержинский и Калининский":
                    sel = @"((district.district = 'Калининский') OR (district.district = 'Дзержинский'))";
                    break;
                case "Кировский и Ленинский":
                    sel = @"((district.district = 'Кировский') OR (district.district = 'Ленинский'))";
                    break;
                case "Октябрьский р-он и центральный административный округ":
                    sel = @"((district.district = 'Заельцовский') OR (district.district = 'Центральный') OR 
                            (district.district = 'Железнодорожный') OR (district.district = 'Октябрьский'))";
                    break;
                case "Советский и Первомайский":
                    sel = @"((district.district = 'Советский') OR (district.district = 'Первомайский'))";
                    break;
                case "Новосибирский р-н, г. Обь и р.п. Кольцово":
                    sel = @"((district.district = 'Новосибирский') OR (district.district = 'г.Обь'))";
                    break;
                case "Барабинский и Здвинский":
                    sel = @"((district.district = 'Барабинский') OR (district.district = 'Здвинский'))";
                    break;
                case "Бердский":
                    sel = @"(district.district = 'Бердский')";
                    break;
                case "Искитимский":
                    sel = @"(district.district = 'Искитимский')";
                    break;
                case "Купинский":
                    sel = @"(district.district = 'Купинский')";
                    break;
                case "Мошковский":
                    sel = @"(district.district = 'Мошковский')";
                    break;
                case "Ордынский":
                    sel = @"(district.district = 'Ордынский')";
                    break;
                case "Сузунский":
                    sel = @"(district.district = 'Сузунский')";
                    break;
                case "Чулымский":
                    sel = @"(district.district = 'Чулымский')";
                    break;
                case "Карасукский и Баганский":
                    sel = @"((district.district = 'Карасукский') OR (district.district = 'Баганский'))";
                    break;
                case "Каргатский и Убинский":
                    sel = @"((district.district = 'Каргатский') OR (district.district = 'Убинский'))";
                    break;
                case "Коченевский и Колыванский":
                    sel = @"((district.district = 'Коченевский') OR (district.district = 'Колыванский'))";
                    break;
                case "Краснозерский, Доволенский и Кочковский":
                    sel = @"((district.district = 'Красноозерский') OR (district.district = 'Доволенский') OR (district.district = 'Кочковский'))";
                    break;
                case "Куйбышевский и Северный":
                    sel = @"((district.district = 'Куйбышевский') OR (district.district = 'Северный'))";
                    break;
                case "Татарский, Усть-Таркский и Чистоозерный":
                    sel = @"((district.district = 'Татарский') OR (district.district = 'Усть-Таркский') OR (district.district = 'Чистоозерный'))";
                    break;
                case "Тогучинский и Болотнинский":
                    sel = @"((district.district = 'Тогучинский') OR (district.district = 'Болотнинский'))";
                    break;
                case "Чановский, Венгеровский и Кыштовский":
                    sel = @"((district.district = 'Чановский') OR (district.district = 'Венгеровский') OR (district.district = 'Кыштовский'))";
                    break;
                case "Черепановский и Маслянинский":
                    sel = @"((district.district = 'Черепановский') OR (district.district = 'Маслянинский'))";
                    break;
            }
            return sel;
        }

        private void PrintForm8_Load(object sender, EventArgs e)
        {
            LoadVoenkomat();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void printButton_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            birthTextBox.Enabled = false;
            districtComboBox.Enabled = false;
            printButton.Enabled = false;
            closeButton.Enabled = false;
            backgroundWorker1.RunWorkerAsync(Tuple.Create(birthTextBox.Text, districtComboBox.Text));
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var t = e.Argument as Tuple<string, string>;
            if (String.IsNullOrWhiteSpace(t.Item1))//вставить проверку, что указаны годы
            {
                MessageBox.Show("Не указан год рождения!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (t.Item2 == "")
            {
                MessageBox.Show("Военкомат не задан!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //здесь формирование ворд-документа
            string select = @"SELECT Student.name, Student.birth, Student.ciizenship, Student.passpSeries, Student.passpNumber, 
                                     [Group].groupName, Student.street, Student.house, Student.flat, Student.phone, sex.sex, district.district, Student.homePhone
                              FROM Student INNER JOIN
                                   [Group] ON Student.idGroup = [Group].id INNER JOIN
                                   sex ON Student.id_sex = sex.id INNER JOIN
                                   district ON Student.id_district = district.id
                              WHERE (Student.birth LIKE '%" + t.Item1 + @"%') AND 
                                    (sex.sex = 'муж.') AND 
                                    ((prikazNumKval = '') AND (prikazNumOut = '')) AND ";
            select += ConvertDistrictNameToDistrictQuery(t.Item2);
            select += @" ORDER BY name";

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

            voenkomat = ConvertVoencomatNameToFullName(t.Item2);

            Word.Application wdApp = new Word.Application();
            Word.Document wdDoc = new Word.Document();
            Object wdMiss = System.Reflection.Missing.Value;

            wdDoc = wdApp.Documents.Add(ref wdMiss, ref wdMiss, ref wdMiss, ref wdMiss);
           // wdApp.Visible = true; //сначала формируется документ, показывать потом
            wdDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            wdDoc.PageSetup.LeftMargin = 40;
            wdDoc.PageSetup.RightMargin = 10;
            wdDoc.PageSetup.TopMargin = 20;
            wdDoc.PageSetup.BottomMargin = 20;
            Word.Table tb;
            Word.Range _range;

            int columnsCount = 6;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, ds1.Tables[0].Rows.Count + 6, columnsCount);
            tb.Columns[1].Width = 40;
            tb.Columns[2].Width = 200;
            tb.Columns[3].Width = 110;
            tb.Columns[4].Width = 150;
            tb.Columns[5].Width = 180;
            tb.Columns[6].Width = 110;

            tb.Rows[1].Height = 40;
            tb.Rows[2].Height = 70;
            tb.Rows[4].Height = 30;
            tb.Rows[5].Height = 40;

            Word.Row row = tb.Rows[1];
            Word.Cell firstCell = row.Cells[1];
            foreach (Word.Cell currCell in row.Cells)
            {
                if (currCell.ColumnIndex != firstCell.ColumnIndex)
                {
                    firstCell.Merge(currCell);
                }
            }
            row = tb.Rows[3];
            firstCell = row.Cells[1];
            foreach (Word.Cell currCell in row.Cells)
            {
                if (currCell.ColumnIndex != firstCell.ColumnIndex)
                {
                    firstCell.Merge(currCell);
                }
            }
            row = tb.Rows[4];
            firstCell = row.Cells[1];
            foreach (Word.Cell currCell in row.Cells)
            {
                if (currCell.ColumnIndex != firstCell.ColumnIndex)
                {
                    firstCell.Merge(currCell);
                }
            }
            row = tb.Rows[5];
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
            wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 0;

            tb.Cell(1, 1).Select();
            //wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            tb.Cell(1, 1).Range.Text = "Приложение №3 \n к инструкции (п.п.9,20) \n Калининский район";

            tb.Rows[3].Select();
            //wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(3, 1).Range.Text = "СПИСОК";
            tb.Rows[4].Select();
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
            tb.Cell(4, 1).Range.Text = "Граждан " + t.Item1 + " года рождения, зарегистрированных и проживающих на территории " + voenkomat;

            tb.Rows[5].Select();
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
            tb.Cell(5, 1).Range.Text = @" ";
            tb.Rows[6].Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            tb.Cell(6, 1).Range.Text = "№ п/п";
            tb.Cell(6, 2).Range.Text = "Фамилия, имя, отчество";
            tb.Cell(6, 3).Range.Text = "Гражданство \n серия и номер паспорта";
            tb.Cell(6, 4).Range.Text = "Место работы (учебы) и занимаемая должность (курс, класс)";
            tb.Cell(6, 5).Range.Text = "Зарегистрированние место жительства, номер телефона (если проживает по другому адресу, указывается место проживания, номер телефона)";
            tb.Cell(6, 6).Range.Text = "Отметка военного комиссариата. За каким порядковым номером учтен в сводном списке.";


            int rowCount = 6;
            int rowNumber = 0;

            foreach (DataRow str in ds1.Tables[0].Rows)
            {
                rowCount++;
                rowNumber++;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.ItemArray[0].ToString() + "\n" + DateTime.Parse(str.ItemArray[1].ToString()).ToShortDateString();
                tb.Cell(rowCount, 3).Range.Text = str.ItemArray[2].ToString() + "\n" + str.ItemArray[3].ToString() + " " + str.ItemArray[4].ToString();
                //tb.Cell(rowCount, 3).Range.Text = DateTime.Parse(str.ItemArray[1].ToString()).ToShortDateString(); //номер курса считать
                tb.Cell(rowCount, 4).Range.Text = "  \n" + Util.CalcKurs(str.ItemArray[5].ToString()) + " курс, уч-ся гр. " + str.ItemArray[5].ToString();
                tb.Cell(rowCount, 5).Range.Text = "ул. " + str.ItemArray[6].ToString() + ", " + str.ItemArray[7].ToString() + " кв. " + str.ItemArray[8].ToString() + "\n"
                                                  + str.ItemArray[11].ToString() + " р-н \n Тел. " + str.ItemArray[9].ToString() + " \n Дом. " + str.ItemArray[12].ToString();
            }

            _range = wdDoc.Range(tb.Cell(7, 2).Range.Start, tb.Cell(rowCount, columnsCount).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
            _range = wdDoc.Range(tb.Cell(7, 3).Range.Start, tb.Cell(rowCount, 3).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            _range = wdDoc.Range(tb.Cell(6, 1).Range.Start, tb.Cell(rowCount, columnsCount).Range.End);
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
            tb2 = wdDoc.Tables.Add(wdApp.Selection.Range, 3, 4);

            tb2.Columns[1].Width = 140;
            tb2.Columns[2].Width = 120;
            tb2.Columns[3].Width = 380;
            tb2.Columns[4].Width = 100;

            tb2.Rows[1].Height = 30;
            tb2.Rows[2].Height = 30;

            tb2.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wdApp.Selection.Range.Font.Name = "Times New Roman";
            wdApp.Selection.Range.Font.Size = 12;
            wdApp.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 0;

            _range = wdDoc.Range(tb2.Cell(3, 2).Range.Start, tb2.Cell(3, 4).Range.End);
            _range.Select();
            wdApp.Selection.Cells.Merge();

            tb2.Cell(1, 2).Range.Text = "Директор колледжа";
            tb2.Cell(1, 4).Range.Text = " ";

            //tb2.Cell(2, 2).Range.Text = "«" + DateTime.Now.Day.ToString() + "»" + DateTime.Now.Month.ToString("MMMM") + " " + DateTime.Now.Year.ToString() + " г.";
            tb2.Cell(2, 2).Range.Text = DateTime.Now.ToLongDateString().ToString();
            tb2.Cell(3, 1).Range.Text = "    М.П.";
            tb2.Cell(3, 2).Range.Text = "Должностное лицо, отвечающее за ведение воинского учёта   контактный телефон  ";

            ////нумерация страниц
            //Word.Window activeWindow = wdDoc.Application.ActiveWindow;
            //object currentPage = Word.WdFieldType.wdFieldPage;
            //object totalPages = Word.WdFieldType.wdFieldNumPages;
            ////переход к редактированию футера
            //activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            //activeWindow.ActivePane.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            ////напечатает номер страницы в формате X стр. из Y
            //activeWindow.Selection.Fields.Add(activeWindow.Selection.Range, ref currentPage, ref wdMiss, ref wdMiss);
            //activeWindow.Selection.TypeText(" стр. из ");
            //activeWindow.Selection.Fields.Add(activeWindow.Selection.Range, ref totalPages, ref wdMiss, ref wdMiss);
            ////выход из футера
            //activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            wdApp.Visible = true; //показать документ пользователю
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Visible = false;
            birthTextBox.Enabled = true;
            districtComboBox.Enabled = true;
            printButton.Enabled = true;
            closeButton.Enabled = true;
        }

    }
}
