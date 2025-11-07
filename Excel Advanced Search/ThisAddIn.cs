using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;

namespace Excel_Advanced_Search
{
    public partial class ThisAddIn
    {
        private UserControl1 userControl;
        public Microsoft.Office.Tools.CustomTaskPane customTaskPane;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Initialize the custom pane
            userControl = new UserControl1();
            customTaskPane = CustomTaskPanes.Add(userControl, "Advanced Search");

            // Position and show by default
            customTaskPane.DockPosition =
                Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            //customTaskPane.Visible = true;

            // Subscribe to Excel events
            Application.SheetSelectionChange += OnSheetSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.SheetSelectionChange -= OnSheetSelectionChange;
        }

        /// <summary>
        /// Toggles the visibility of the Kouach Search custom task pane.
        /// Called from the Ribbon button.
        /// </summary>
        public void ToggleSearchPane()
        {
            if (customTaskPane == null)
                return;

            customTaskPane.Visible = !customTaskPane.Visible;
        }

        /// <summary>
        /// Updates the active textbox when a range is selected in Excel.
        /// </summary>
        private void OnSheetSelectionChange(object sheet, Excel.Range target)
        {
            try
            {
                if (userControl?.ActiveControl == null ||
                    !userControl.CanFocus ||
                    userControl.ActiveControl.Name.Equals("textBox3", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                Excel.Range selectedRange = Application.Selection as Excel.Range;
                if (selectedRange == null)
                    return;

                string sheetName = selectedRange.Worksheet?.Name ?? string.Empty;
                string cellAddress = selectedRange.Address?.Replace("$", string.Empty) ?? string.Empty;

                userControl.ActiveControl.Text = $"{sheetName}!{cellAddress}";

                Debug.WriteLine($"Active cell changed to: {sheetName}!{cellAddress}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Error] OnSheetSelectionChange: {ex.Message}");
            }
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
