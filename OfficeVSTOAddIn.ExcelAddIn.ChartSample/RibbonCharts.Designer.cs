namespace ExcelAddIn.ChartSample
{
    partial class RibbonCharts : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonCharts()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupChartSample = this.Factory.CreateRibbonGroup();
            this.buttonAddChart = this.Factory.CreateRibbonButton();
            this.buttonAddCharts = this.Factory.CreateRibbonButton();
            this.buttonSetData = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupChartSample.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupChartSample);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupChartSample
            // 
            this.groupChartSample.Items.Add(this.buttonSetData);
            this.groupChartSample.Items.Add(this.buttonAddChart);
            this.groupChartSample.Items.Add(this.buttonAddCharts);
            this.groupChartSample.Label = "Chart Sample";
            this.groupChartSample.Name = "groupChartSample";
            // 
            // buttonAddChart
            // 
            this.buttonAddChart.Label = "Add One Chart";
            this.buttonAddChart.Name = "buttonAddChart";
            this.buttonAddChart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddChart_Click);
            // 
            // buttonAddCharts
            // 
            this.buttonAddCharts.Label = "Add Chats(Randomly)";
            this.buttonAddCharts.Name = "buttonAddCharts";
            this.buttonAddCharts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddCharts_Click);
            // 
            // buttonSetData
            // 
            this.buttonSetData.Label = "Set Sample Data";
            this.buttonSetData.Name = "buttonSetData";
            this.buttonSetData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSetData_Click);
            // 
            // RibbonCharts
            // 
            this.Name = "RibbonCharts";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonCharts_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupChartSample.ResumeLayout(false);
            this.groupChartSample.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupChartSample;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddChart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddCharts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSetData;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonCharts Ribbon1
        {
            get { return this.GetRibbon<RibbonCharts>(); }
        }
    }
}
