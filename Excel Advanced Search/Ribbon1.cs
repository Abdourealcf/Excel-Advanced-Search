using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Advanced_Search
{
    public partial class Ribbon1
    {

        public Ribbon1()
    : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }


        private void TogglePaneButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Globals.ThisAddIn?.ToggleSearchPane(Globals.ThisAddIn.Panes[activeWorkbook]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"An error occurred while toggling the search pane:\n{ex.Message}",
                    "Advanced Search",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

    }
}
