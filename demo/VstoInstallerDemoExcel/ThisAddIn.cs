using System.Windows.Forms;

namespace VstoInstallerDemoExcel
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (Properties.Settings.Default.FirstRun)
            {
                Properties.Settings.Default.FirstRun = false;
                Properties.Settings.Default.Save();
                MessageBox.Show(
                    "The VstoAddinInstaller demo add-in for Excel was successfully installed.\n" +
                    "This message will not be shown again.\n" +
                    "You can uninstall the add-in now.",
                    "VstoAddinInstaller Demo Add-in",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
