using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace Excel_Advanced_Search
{
    public partial class UserControl1 : UserControl
    {


        public UserControl1()
        {
            InitializeComponent();
        }

        private async void TextBox3_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text) ||
        string.IsNullOrWhiteSpace(textBox2.Text) ||
        string.IsNullOrWhiteSpace(textBox4.Text))
                return;

            await Task.Run(() =>
            {
                this.Invoke(new System.Action(() =>
                {
                    try
                    {
                        ProcessExcelData();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Excel operation failed: {ex.Message}",
                            "Excel Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }));
            });
        }

        private void ProcessExcelData()
        {
            var app = Globals.ThisAddIn.Application;

            Worksheet sourceSheet = app.Sheets[Helper.GetSheetName(textBox1.Text)];
            Worksheet destinationSheet = app.Sheets[Helper.GetSheetName(textBox4.Text)];

            Range tableRange = sourceSheet.Range[Helper.GetRangeAddress(textBox1.Text)];
            Range columnRange = sourceSheet.Range[Helper.GetRangeAddress(textBox2.Text)];
            Range outputStartRange = destinationSheet.Range[Helper.GetRangeAddress(textBox4.Text)];

            if (columnRange.Columns.Count != 1 || tableRange.Columns.Count < 1)
                return;

            var (startLetter, endLetter) = Helper.GetRangeLetters(tableRange);
            var (outputStartLetter, outputEndLetter) = Helper.GetRangeLetters(outputStartRange);
            int outputRow = Helper.GetRowNumber(outputStartRange.Address);
            int colCount = tableRange.Columns.Count;


            int lastRow = destinationSheet.Rows.Count;

            Range fullOutputRange = destinationSheet.Range[
                destinationSheet.Cells[outputRow, outputStartRange.Column],
                destinationSheet.Cells[lastRow, outputStartRange.Column + colCount - 1]
            ];
            fullOutputRange.ClearContents();

            if (string.IsNullOrWhiteSpace(textBox3.Text))
                return;

            destinationSheet.Activate();

            string[] searchTerms = textBox3.Text.Split('|').Length > 1 ? textBox3.Text.Split('|') : new string[] { textBox3.Text };

            object[,] columnValues = columnRange.Value2 as object[,];
            object[,] tableValues = tableRange.Value2 as object[,];
            int rowCount = columnRange.Rows.Count;

            var matchedRows = new List<object[]>();

            for (int i = 1; i <= rowCount; i++)
            {
                string cellText = columnValues[i, 1]?.ToString();
                if (string.IsNullOrWhiteSpace(cellText)) continue;

                for (int indx = 0; indx < searchTerms.Length; indx++)
                {
                    if (string.IsNullOrWhiteSpace(searchTerms[indx])) continue;
                    if (CultureInfo.CurrentCulture.CompareInfo
                                        .IndexOf(cellText, searchTerms[indx], CompareOptions.IgnoreCase) >= 0)
                    {
                        var row = new object[colCount];
                        for (int col = 1; col <= colCount; col++)
                            row[col - 1] = tableValues[i, col];
                        matchedRows.Add(row);
                        break;//Stop checking other keywords — the row is already in matchedRows, so additional matches would be redundant.
                    }
                }


            }


            if (matchedRows.Count > 0)
            {
                object[,] outputArray = new object[matchedRows.Count, colCount];
                for (int r = 0; r < matchedRows.Count; r++)
                    for (int c = 0; c < colCount; c++)
                        outputArray[r, c] = matchedRows[r][c];
                Range outputRange = destinationSheet.Range[$"{outputStartLetter}{outputRow}"]
                    .Resize[matchedRows.Count, colCount];
                outputRange.Value2 = outputArray;
            }
        }

        #region Utility Methods



        private static void ClearOutputRange(Worksheet sheet, string startLetter, string endLetter, int startRow)
        {
            int startCol = sheet.Range[$"{startLetter}{startRow}"].Column;
            int endCol = sheet.Range[$"{endLetter}{startRow}"].Column;

            for (int col = startCol; col <= endCol; col++)
            {
                Range startCell = sheet.Cells[startRow, col];
                Range lastCell = startCell.End[XlDirection.xlDown];
                if (lastCell == null) continue;

                Range cleanRange = sheet.Range[$"{startLetter}{startRow}:{endLetter}{lastCell.Row}"];
                cleanRange.ClearContents();
            }
        }



        private void TextBox_Enter(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox TextBox = (System.Windows.Forms.TextBox)sender;

            if (string.IsNullOrWhiteSpace(TextBox.Text)) return;
            SelectExcelRange(TextBox.Text);
        }

        public static void SelectExcelRange(string rangeRef)
        {
            if (string.IsNullOrWhiteSpace(rangeRef)) return;
            //sleep for 50ms to prevent mutiple fast selection by user
            System.Threading.Thread.Sleep(50);
            var app = Globals.ThisAddIn.Application;
            Worksheet sheet = app.Sheets[Helper.GetSheetName(rangeRef)];
            Range range = sheet.Range[Helper.GetRangeAddress(rangeRef)];
            sheet.Activate();
            range?.Select();
        }

        #endregion

        private void textBox_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox txtbx = (System.Windows.Forms.TextBox)sender;

            string message = (txtbx.Name == "textBox1")
            ? "Select the range of the table that contains the data to be searched."
            : (txtbx.Name == "textBox2")
            ? "Select the column range where you want to search for the keywords."
            : (txtbx.Name == "textBox3")
            ? "Enter one or more keywords, separated by the '|' character."
            : "Select the cell where the search results should be displayed.";

            toolTip1.Show(message, txtbx);
        }

       

    }
}
