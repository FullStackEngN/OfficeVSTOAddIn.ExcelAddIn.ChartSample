/// <summary>
/// WARNING: ANY USE BY YOU OF THE SAMPLE CODE PROVIDED IN THIS FILE IS AT YOUR OWN RISK. 
/// Microsoft provides this code "as is" without warranty of any kind, either express or implied, 
/// including but not limited to the implied warranties of merchantability and/or fitness for a particular purpose.
/// </summary>

namespace ExcelAddIn.ChartSample
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
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
