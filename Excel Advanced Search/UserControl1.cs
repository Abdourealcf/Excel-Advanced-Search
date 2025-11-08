using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Advanced_Search
{
    public partial class UserControl1 : UserControl
    {
        private static readonly Regex LetterRegex = new Regex("[A-Za-z]+", RegexOptions.Compiled);
        private static readonly Regex NumberRegex = new Regex(@"\d+", RegexOptions.Compiled);

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

            // Run the heavy Excel work asynchronously on a background thread,
            // but Excel COM operations must run on the UI thread.
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

        /// <summary>
        /// Executes the Excel processing logic on a background thread,
        /// while ensuring Excel COM calls are invoked on its main thread.
        /// </summary>


        /// <summary>
        /// Main logic that filters, clears, and copies rows between Excel ranges.
        /// </summary>
        private void ProcessExcelData()
        {
            var app = Globals.ThisAddIn.Application;

            Worksheet sourceSheet = app.Sheets[GetSheetName(textBox1.Text)];
            Worksheet destinationSheet = app.Sheets[GetSheetName(textBox4.Text)];

            Range tableRange = sourceSheet.Range[GetRangeAddress(textBox1.Text)];
            Range columnRange = sourceSheet.Range[GetRangeAddress(textBox2.Text)];
            Range outputStartRange = destinationSheet.Range[GetRangeAddress(textBox4.Text)];

            if (columnRange.Columns.Count != 1 || tableRange.Columns.Count < 1)
                return;

            var (startLetter, endLetter) = GetRangeLetters(tableRange);
            var (outputStartLetter, outputEndLetter) = GetRangeLetters(outputStartRange);
            int outputRow = GetRowNumber(outputStartRange.Address);
            int colCount = tableRange.Columns.Count;

            // ✅ Clear any existing contents below the start row

            int lastRow = destinationSheet.Rows.Count;

            // Clear all columns in output area
            Range fullOutputRange = destinationSheet.Range[
                destinationSheet.Cells[outputRow, outputStartRange.Column],
                destinationSheet.Cells[lastRow, outputStartRange.Column + colCount - 1]
            ];
            fullOutputRange.ClearContents();

            if (string.IsNullOrWhiteSpace(textBox3.Text))
                return;

            destinationSheet.Activate();

            string[] searchTerms = textBox3.Text.Split('|').Length > 1 ? textBox3.Text.Split('|') : new string[] { textBox3.Text };

            // Load source ranges into arrays
            object[,] columnValues = columnRange.Value2 as object[,];
            object[,] tableValues = tableRange.Value2 as object[,];
            int rowCount = columnRange.Rows.Count;

            var matchedRows = new List<object[]>();

            for (int i = 1; i <= rowCount; i++)
            {
                string cellText = columnValues[i, 1]?.ToString();
                if (string.IsNullOrWhiteSpace(cellText)) continue;

                foreach (string searchTerm in searchTerms)
                {
                    if (CultureInfo.CurrentCulture.CompareInfo
                                        .IndexOf(cellText, searchTerm, CompareOptions.IgnoreCase) >= 0)
                    {
                        var row = new object[colCount];
                        for (int col = 1; col <= colCount; col++)
                            row[col - 1] = tableValues[i, col];

                        matchedRows.Add(row);
                    }
                }


            }

            // ✅ Write all matched rows at once
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

        private static (string startLetter, string endLetter) GetRangeLetters(Range range)
        {
            string[] parts = range.Address.Split(':');
            return (ExtractLetters(parts[0]), parts.Length > 1 ? ExtractLetters(parts[1]) : ExtractLetters(parts[0]));
        }

        private static string ExtractLetters(string address) =>
            LetterRegex.Match(address).Value;

        private static int GetRowNumber(string address, int index = 0)
        {
            string[] parts = address.Split(':');
            return int.TryParse(NumberRegex.Match(parts[index]).Value, out int num) ? num : 1;
        }

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

        private static string GetSheetName(string fullRef) =>
            fullRef.Split('!')[0];

        private static string GetRangeAddress(string fullRef) =>
            fullRef.Split('!')[1];

        private void TextBox_Enter(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox TextBox = (System.Windows.Forms.TextBox)sender;

            if (string.IsNullOrWhiteSpace(TextBox.Text)) return;
            SelectExcelRange(TextBox.Text);
        }

        private static void SelectExcelRange(string rangeRef)
        {
            var app = Globals.ThisAddIn.Application;
            Worksheet sheet = app.Sheets[GetSheetName(rangeRef)];
            Range range = sheet.Range[GetRangeAddress(rangeRef)];

            sheet.Activate();
            range?.Select();
        }

        #endregion
    }
}
