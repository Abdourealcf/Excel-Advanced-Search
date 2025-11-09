using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Microsoft.Office.Tools;

namespace Excel_Advanced_Search
{
    public partial class ThisAddIn
    {

        public Dictionary<Excel.Workbook, CustomTaskPane> Panes = new Dictionary<Excel.Workbook, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {

            Application.WorkbookActivate += Application_WorkbookActivate;
            Application.SheetSelectionChange += OnSheetSelectionChange;

            if (Application.ActiveWorkbook != null) AddCustomTaskPaneForWorkbook(Application.ActiveWorkbook);

        }


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.SheetSelectionChange -= OnSheetSelectionChange;
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Panes.Clear();
        }

        public void ToggleSearchPane(CustomTaskPane customTaskPane)
        {
            if (customTaskPane != null) customTaskPane.Visible = !customTaskPane.Visible;
        }

        private void OnSheetSelectionChange(object sheet, Excel.Range target)
        {
            //sleep for 50ms to prevent mutiple fast selection by user
            System.Threading.Thread.Sleep(50);

            UserControl1 myControl = Globals.ThisAddIn.Panes[Application.ActiveWorkbook].Control as UserControl1;
            try
            {
                if (myControl?.ActiveControl == null ||
                    !myControl.CanFocus ||
                    myControl.ActiveControl.Name.Equals("textBox3", StringComparison.OrdinalIgnoreCase)) return;


                Excel.Range selectedRange = Application.Selection as Excel.Range;

                if (selectedRange == null) return;

                string sheetName = selectedRange.Worksheet?.Name ?? string.Empty;
                string cellAddress = selectedRange.Address?.Replace("$", string.Empty) ?? string.Empty;

                if (String.IsNullOrEmpty(cellAddress)) return;//just incase

                //return if selection Address is the same as the text in textbox to prevent infinit loop when Forcing column range to match the table range vertically
                if (myControl.ActiveControl.Text == $"{sheetName}!{cellAddress}") return;

                if (myControl.ActiveControl == myControl.Controls["textBox2"])
                {
                    // Force column range to match the table range vertically.

                    string textbox1StartNumber = Helper.GetRowNumber(Helper.GetRangeAddress(myControl.Controls["textBox1"].Text), 0).ToString();
                    string textbox1EndNumber = Helper.GetRowNumber(Helper.GetRangeAddress(myControl.Controls["textBox1"].Text), 1).ToString();


                        Regex letterRegex = new Regex("[a-zA-Z]+");
                        string textbox2letter = letterRegex.Match(cellAddress).Value;

                        cellAddress = $"{textbox2letter}{textbox1StartNumber}:{textbox2letter}{textbox1EndNumber}";
                    

                    myControl.ActiveControl.Text = $"{sheetName}!{cellAddress}";
                    //force the TextBox_Enter event for textbox2 to update the new column range visually.
                    UserControl1.TextBox_Enter(myControl.Controls["textBox2"],  EventArgs.Empty);
                    return;
                }

                myControl.ActiveControl.Text = $"{sheetName}!{cellAddress}";

                Debug.WriteLine($"Active cell changed to: {sheetName}!{cellAddress}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] OnSheetSelectionChange: {ex.Message}");
            }
        }



        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            if (!Panes.ContainsKey(workbook)) AddCustomTaskPaneForWorkbook(workbook);

        }


        private void AddCustomTaskPaneForWorkbook(Excel.Workbook workbook)
        {
            UserControl1 myControl = new UserControl1();
            CustomTaskPane customTaskPane = CustomTaskPanes.Add(myControl, "Advanced Search", workbook.Windows[1]);
            customTaskPane.DockPosition =
                 Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Panes.Add(workbook, customTaskPane);
        }


        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
