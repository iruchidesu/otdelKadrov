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

namespace WindowsFormsApplication1
{
    class PrintDataGrid
    {
        DataGridViewRowCollection rowColl;

        public PrintDataGrid(DataGridViewRowCollection rowCollection)
        {
            this.rowColl = rowCollection;
        }

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

        public void Print()
        {
            Word.Application wdApp = new Word.Application();
            Word.Document wdDoc = new Word.Document();
            Object wdMiss = System.Reflection.Missing.Value;

            wdDoc = wdApp.Documents.Add(ref wdMiss, ref wdMiss, ref wdMiss, ref wdMiss);
            wdApp.Visible = true; //сначала формируется документ, показывать потом
            wdDoc.PageSetup.LeftMargin = 40;
            wdDoc.PageSetup.RightMargin = 25;
            wdDoc.PageSetup.TopMargin = 20;
            wdDoc.PageSetup.BottomMargin = 20;
            Word.Table tb;
            Word.Range _range;

            int columnsCount = 4;

            tb = wdDoc.Tables.Add(wdApp.Selection.Range, rowColl.Count + 2, columnsCount);
            tb.Columns[1].Width = 40;
            tb.Rows[1].Height = 30;
            tb.Columns[2].Width = 240;
            tb.Rows[2].Height = 30;
            tb.Columns[3].Width = 80;
            tb.Rows[3].Height = 40;
            tb.Columns[4].Width = 180;

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
            //tb.Cell(2, 1).Range.Text = "студентов, обучающихся на коммерческой основе на " + DateTime.Now.ToShortDateString() + " г.";

            tb.Rows[3].Select();
            wdApp.Selection.Font.Bold = 1;
            tb.Cell(3, 1).Range.Text = "№ п/п";
            tb.Cell(3, 2).Range.Text = "ФИО";
            tb.Cell(3, 3).Range.Text = "Дата рождения";
            tb.Cell(3, 4).Range.Text = "Примечание";


            int rowCount = 3;
            int rowNumber = 0;

            foreach (DataGridViewRow str in rowColl)
            {
                rowCount++;
                rowNumber++;
                if (rowNumber == rowColl.Count)
                    break;
                tb.Cell(rowCount, 1).Range.Text = rowNumber.ToString() + ".";
                tb.Cell(rowCount, 2).Range.Text = str.Cells[0].Value.ToString(); //фио
                tb.Cell(rowCount, 3).Range.Text = str.Cells[1].Value.ToString().Remove(11); //дата рождения
                tb.Cell(rowCount, 4).Range.Text = str.Cells[30].Value.ToString(); //примечание
            }

            _range = wdDoc.Range(tb.Cell(4, 2).Range.Start, tb.Cell(rowCount, 2).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 10;
            _range = wdDoc.Range(tb.Cell(4, 4).Range.Start, tb.Cell(rowCount, 4).Range.End);
            _range.Select();
            wdApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            wdApp.Selection.ParagraphFormat.SpaceAfter = 10;
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

            //wdApp.Visible = true; //показать документ пользователю
        }

        public void PrintExcel()
        {
            Object exMiss = System.Reflection.Missing.Value;
            Excel.Workbook exclBook;
            Excel.Worksheet exclSheet;
            Excel.Application exclApp = new Excel.Application();
            
            exclBook = exclApp.Workbooks.Add();
            exclSheet = (Excel.Worksheet)exclBook.Sheets[1];

            //exclApp.Visible = true;

            exclSheet.Cells[1, 1] = "№ п/п";
            exclSheet.Cells[1, 2] = "ФИО";
            exclSheet.Cells[1, 3] = "Дата рождения";
            exclSheet.Cells[1, 4] = "Пол";
            exclSheet.Cells[1, 5] = "В/о";
            exclSheet.Cells[1, 6] = "Категория годности";
            exclSheet.Cells[1, 7] = "Индекс";
            exclSheet.Cells[1, 8] = "Город (область)";
            exclSheet.Cells[1, 9] = "Район";
            exclSheet.Cells[1, 10] = "Населенный пункт";
            exclSheet.Cells[1, 11] = "Улица";
            exclSheet.Cells[1, 12] = "Дом";
            exclSheet.Cells[1, 13] = "Квартира";
            exclSheet.Cells[1, 14] = "Телефон";
            exclSheet.Cells[1, 15] = "Дом.телефон";
            exclSheet.Cells[1, 16] = "Гражданство";
            exclSheet.Cells[1, 17] = "Паспорт серия";
            exclSheet.Cells[1, 18] = "Паспорт номер";
            exclSheet.Cells[1, 19] = "Кем выдан";
            exclSheet.Cells[1, 20] = "Когда выдан";
            exclSheet.Cells[1, 21] = "Группа";
            exclSheet.Cells[1, 22] = "Отделение";
            exclSheet.Cells[1, 23] = "Бюд./Ком.";
            exclSheet.Cells[1, 24] = "№ приказа о зачислении";
            exclSheet.Cells[1, 25] = "Дата зачисления";
            exclSheet.Cells[1, 26] = "№ приказа о отчислении";
            exclSheet.Cells[1, 27] = "Дата отчисления";
            exclSheet.Cells[1, 28] = "Причина отчисления";
            exclSheet.Cells[1, 29] = "Квалификация";
            exclSheet.Cells[1, 30] = "№ приказа о квалификации";
            exclSheet.Cells[1, 31] = "Академический отпуск";
            exclSheet.Cells[1, 32] = "Примечание";

            int rowCount = 1;
            int rowNumber = 0;

            foreach (DataGridViewRow str in rowColl)
            {
                rowCount++;
                rowNumber++;
                if (rowNumber == rowColl.Count)
                    break;
                exclSheet.Cells[rowCount, 1] = rowNumber.ToString() + ".";
                exclSheet.Cells[rowCount, 2] = str.Cells[0].Value.ToString();
                if (str.Cells[1].Value.ToString().Remove(11) != "01.01.1900 ")
                    exclSheet.Cells[rowCount, 3] = str.Cells[1].Value.ToString().Remove(11);
                else
                    exclSheet.Cells[rowCount, 3] = "";
                exclSheet.Cells[rowCount, 4] = str.Cells[2].Value.ToString();
                exclSheet.Cells[rowCount, 5] = str.Cells[3].Value.ToString();
                exclSheet.Cells[rowCount, 6] = str.Cells[4].Value.ToString();
                exclSheet.Cells[rowCount, 7] = str.Cells[5].Value.ToString();
                exclSheet.Cells[rowCount, 8] = str.Cells[6].Value.ToString();
                exclSheet.Cells[rowCount, 9] = str.Cells[7].Value.ToString();
                exclSheet.Cells[rowCount, 10] = str.Cells[8].Value.ToString();
                exclSheet.Cells[rowCount, 11] = str.Cells[9].Value.ToString();
                exclSheet.Cells[rowCount, 12] = str.Cells[10].Value.ToString();
                exclSheet.Cells[rowCount, 13] = str.Cells[11].Value.ToString();
                exclSheet.Cells[rowCount, 14] = str.Cells[12].Value.ToString();
                exclSheet.Cells[rowCount, 15] = str.Cells[13].Value.ToString();
                exclSheet.Cells[rowCount, 16] = str.Cells[14].Value.ToString();
                exclSheet.Cells[rowCount, 17] = str.Cells[15].Value.ToString();
                exclSheet.Cells[rowCount, 18] = str.Cells[16].Value.ToString();
                exclSheet.Cells[rowCount, 19] = str.Cells[17].Value.ToString();
                if (str.Cells[18].Value.ToString().Remove(11) != "01.01.1900 ")
                    exclSheet.Cells[rowCount, 20] = str.Cells[18].Value.ToString().Remove(11);
                else
                    exclSheet.Cells[rowCount, 20] = "";
                exclSheet.Cells[rowCount, 21] = str.Cells[19].Value.ToString();
                exclSheet.Cells[rowCount, 22] = str.Cells[20].Value.ToString();
                exclSheet.Cells[rowCount, 23] = str.Cells[21].Value.ToString();
                exclSheet.Cells[rowCount, 24] = str.Cells[22].Value.ToString();
                if (str.Cells[23].Value.ToString().Remove(11) != "01.01.1900 ")
                    exclSheet.Cells[rowCount, 25] = str.Cells[23].Value.ToString().Remove(11);
                else
                    exclSheet.Cells[rowCount, 25] = "";
                exclSheet.Cells[rowCount, 26] = str.Cells[24].Value.ToString();
                if (str.Cells[25].Value.ToString().Remove(11) != "01.01.1900 ")
                    exclSheet.Cells[rowCount, 27] = str.Cells[25].Value.ToString().Remove(11);
                else
                    exclSheet.Cells[rowCount, 27] = "";
                exclSheet.Cells[rowCount, 28] = str.Cells[26].Value.ToString();
                exclSheet.Cells[rowCount, 29] = str.Cells[27].Value.ToString();
                exclSheet.Cells[rowCount, 30] = str.Cells[28].Value.ToString();
                exclSheet.Cells[rowCount, 31] = str.Cells[29].Value.ToString();
                exclSheet.Cells[rowCount, 32] = str.Cells[30].Value.ToString();
            }

            exclApp.Visible = true;

            ReleaseObject(exclApp);
        }

    }
}
