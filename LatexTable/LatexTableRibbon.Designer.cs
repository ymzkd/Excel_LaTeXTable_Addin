namespace LatexTable
{
    partial class LatexTableRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public LatexTableRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.LatexTable_tab = this.Factory.CreateRibbonTab();
            this.createTable = this.Factory.CreateRibbonGroup();
            this.Convert = this.Factory.CreateRibbonButton();
            this.SaveFileButton = this.Factory.CreateRibbonButton();
            this.settings = this.Factory.CreateRibbonGroup();
            this.enableCentering = this.Factory.CreateRibbonCheckBox();
            this.fitWidth = this.Factory.CreateRibbonCheckBox();
            this.position = this.Factory.CreateRibbonEditBox();
            this.skipHidden = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.hasCaption = this.Factory.CreateRibbonCheckBox();
            this.Caption = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.hasLabel = this.Factory.CreateRibbonCheckBox();
            this.Label = this.Factory.CreateRibbonEditBox();
            this.LatexTable_tab.SuspendLayout();
            this.createTable.SuspendLayout();
            this.settings.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // LatexTable_tab
            // 
            this.LatexTable_tab.Groups.Add(this.createTable);
            this.LatexTable_tab.Groups.Add(this.settings);
            this.LatexTable_tab.Groups.Add(this.group1);
            this.LatexTable_tab.Groups.Add(this.group2);
            this.LatexTable_tab.Label = "LatexTable";
            this.LatexTable_tab.Name = "LatexTable_tab";
            // 
            // createTable
            // 
            this.createTable.Items.Add(this.Convert);
            this.createTable.Items.Add(this.SaveFileButton);
            this.createTable.Label = "Create Table";
            this.createTable.Name = "createTable";
            // 
            // Convert
            // 
            this.Convert.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Convert.Description = "Convert selected Range to Latex table";
            this.Convert.Label = "Copy as Table";
            this.Convert.Name = "Convert";
            this.Convert.OfficeImageId = "Copy";
            this.Convert.ShowImage = true;
            this.Convert.SuperTip = "Convert selected Range to Latex table";
            this.Convert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // SaveFileButton
            // 
            this.SaveFileButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SaveFileButton.Label = "Save to File";
            this.SaveFileButton.Name = "SaveFileButton";
            this.SaveFileButton.OfficeImageId = "SaveAndClose";
            this.SaveFileButton.ShowImage = true;
            this.SaveFileButton.SuperTip = "Save LaTex Table text as string";
            this.SaveFileButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveFileButton_Click);
            // 
            // settings
            // 
            this.settings.Items.Add(this.enableCentering);
            this.settings.Items.Add(this.fitWidth);
            this.settings.Items.Add(this.position);
            this.settings.Items.Add(this.skipHidden);
            this.settings.Label = "Settings";
            this.settings.Name = "settings";
            // 
            // enableCentering
            // 
            this.enableCentering.Checked = true;
            this.enableCentering.Description = "Centering Table";
            this.enableCentering.Label = "Enable Centering";
            this.enableCentering.Name = "enableCentering";
            this.enableCentering.ScreenTip = "Centering Table";
            this.enableCentering.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.enableCentering_Click);
            // 
            // fitWidth
            // 
            this.fitWidth.Description = "scale table to text width";
            this.fitWidth.Label = "Enable Width Fit";
            this.fitWidth.Name = "fitWidth";
            this.fitWidth.SuperTip = "scale table to text width";
            this.fitWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fitWidth_Click);
            // 
            // position
            // 
            this.position.Label = "Table Position";
            this.position.MaxLength = 6;
            this.position.Name = "position";
            this.position.SizeString = "WWWWWWW";
            this.position.SuperTip = "Table Position Control String";
            this.position.Text = "htpb";
            this.position.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.position_TextChanged);
            // 
            // skipHidden
            // 
            this.skipHidden.Label = "Skip Hidden";
            this.skipHidden.Name = "skipHidden";
            this.skipHidden.SuperTip = "If checked, hidden cells are not conteined to generated Table";
            this.skipHidden.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.skipHidden_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.hasCaption);
            this.group1.Items.Add(this.Caption);
            this.group1.Label = "Caption";
            this.group1.Name = "group1";
            // 
            // hasCaption
            // 
            this.hasCaption.Label = "Add Caption";
            this.hasCaption.Name = "hasCaption";
            this.hasCaption.ScreenTip = "Add Caption to Table";
            this.hasCaption.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hasCaption_Click);
            // 
            // Caption
            // 
            this.Caption.Label = "Caption";
            this.Caption.Name = "Caption";
            this.Caption.ScreenTip = "Table Caption Text";
            this.Caption.SizeString = "WWWWWWW";
            this.Caption.Text = null;
            // 
            // group2
            // 
            this.group2.Items.Add(this.hasLabel);
            this.group2.Items.Add(this.Label);
            this.group2.Label = "Label";
            this.group2.Name = "group2";
            // 
            // hasLabel
            // 
            this.hasLabel.Description = "Add Label to Table";
            this.hasLabel.Label = "Add Label";
            this.hasLabel.Name = "hasLabel";
            this.hasLabel.ScreenTip = "Add Label to Table";
            this.hasLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hasLabel_Click);
            // 
            // Label
            // 
            this.Label.Label = "Label";
            this.Label.Name = "Label";
            this.Label.ScreenTip = "Table Label Text";
            this.Label.SizeString = "WWWWWWW";
            this.Label.Text = null;
            // 
            // LatexTableRibbon
            // 
            this.Name = "LatexTableRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.LatexTable_tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.LatexTableRibbon_Load);
            this.LatexTable_tab.ResumeLayout(false);
            this.LatexTable_tab.PerformLayout();
            this.createTable.ResumeLayout(false);
            this.createTable.PerformLayout();
            this.settings.ResumeLayout(false);
            this.settings.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab LatexTable_tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup createTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Convert;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup settings;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox enableCentering;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox fitWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox position;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox hasCaption;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Caption;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox hasLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Label;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SaveFileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox skipHidden;
    }

    partial class ThisRibbonCollection
    {
        internal LatexTableRibbon LatexTableRibbon
        {
            get { return this.GetRibbon<LatexTableRibbon>(); }
        }
    }
}
