using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;

namespace Excel_Advanced_Search
{
    public partial class Ribbon1
    {

        public Ribbon1()
    : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateButtonLabel();
        }

        private void TogglePaneButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn?.ToggleSearchPane();
                UpdateButtonLabel();
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

        private void UpdateButtonLabel()
        {
            var pane = Globals.ThisAddIn?.customTaskPane;
            button1.Label = pane?.Visible == true ? "Hide Search Panel" : "Show Search Panel";
        }
    }
}
