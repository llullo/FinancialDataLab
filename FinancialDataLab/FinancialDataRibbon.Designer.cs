namespace FinancialDataLab
{
    partial class FinancialDataRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FinancialDataRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabFinancialDataLab = this.Factory.CreateRibbonTab();
            this.groupDataImport = this.Factory.CreateRibbonGroup();
            this.btnLoadObservations = this.Factory.CreateRibbonButton();
            this.btnPreviousPage = this.Factory.CreateRibbonButton();
            this.btnNextPage = this.Factory.CreateRibbonButton();
            this.tabFinancialDataLab.SuspendLayout();
            this.groupDataImport.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabFinancialDataLab
            // 
            this.tabFinancialDataLab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabFinancialDataLab.Groups.Add(this.groupDataImport);
            this.tabFinancialDataLab.Label = "Financial Data Lab";
            this.tabFinancialDataLab.Name = "tabFinancialDataLab";
            // 
            // groupDataImport
            // 
            this.groupDataImport.Items.Add(this.btnLoadObservations);
            this.groupDataImport.Items.Add(this.btnPreviousPage);
            this.groupDataImport.Items.Add(this.btnNextPage);
            this.groupDataImport.Label = "Data Import";
            this.groupDataImport.Name = "groupDataImport";
            // 
            // btnLoadObservations
            // 
            this.btnLoadObservations.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoadObservations.Label = "Load Observations";
            this.btnLoadObservations.Name = "btnLoadObservations";
            this.btnLoadObservations.ScreenTip = "Load financial observations from SQL Server";
            this.btnLoadObservations.ShowImage = true;
            this.btnLoadObservations.SuperTip = "Retrieves financial data through ADO.NET and writes it into the active Excel work" +
    "sheet.";
            this.btnLoadObservations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadObservations_Click);
            // 
            // btnPreviousPage
            // 
            this.btnPreviousPage.Label = "Previous Page";
            this.btnPreviousPage.Name = "btnPreviousPage";
            this.btnPreviousPage.ScreenTip = "Load the previous page";
            this.btnPreviousPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPreviousPage_Click);
            // 
            // btnNextPage
            // 
            this.btnNextPage.Label = "Next Page";
            this.btnNextPage.Name = "btnNextPage";
            this.btnNextPage.ScreenTip = "Load the next page";
            this.btnNextPage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNextPage_Click);
            // 
            // FinancialDataRibbon
            // 
            this.Name = "FinancialDataRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabFinancialDataLab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FinancialDataRibbon_Load);
            this.tabFinancialDataLab.ResumeLayout(false);
            this.tabFinancialDataLab.PerformLayout();
            this.groupDataImport.ResumeLayout(false);
            this.groupDataImport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabFinancialDataLab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDataImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadObservations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPreviousPage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNextPage;
    }

    partial class ThisRibbonCollection
    {
        internal FinancialDataRibbon FinancialDataRibbon
        {
            get { return this.GetRibbon<FinancialDataRibbon>(); }
        }
    }
}
