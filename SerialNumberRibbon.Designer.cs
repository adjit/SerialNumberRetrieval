﻿namespace SerialNumberRetrieval
{
    partial class SerialNumberRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SerialNumberRibbon()
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
            this.serialNumberGroup = this.Factory.CreateRibbonGroup();
            this.getSerialNumbersButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.serialNumberGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.serialNumberGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // serialNumberGroup
            // 
            this.serialNumberGroup.Items.Add(this.getSerialNumbersButton);
            this.serialNumberGroup.Label = "Serial Number Group";
            this.serialNumberGroup.Name = "serialNumberGroup";
            // 
            // getSerialNumbersButton
            // 
            this.getSerialNumbersButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.getSerialNumbersButton.Image = global::SerialNumberRetrieval.Properties.Resources.tracking_number_512;
            this.getSerialNumbersButton.Label = "Get Serial Numbers";
            this.getSerialNumbersButton.Name = "getSerialNumbersButton";
            this.getSerialNumbersButton.ShowImage = true;
            this.getSerialNumbersButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.getSerialNumbersButton_Click);
            // 
            // SerialNumberRibbon
            // 
            this.Name = "SerialNumberRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.serialNumberGroup.ResumeLayout(false);
            this.serialNumberGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup serialNumberGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton getSerialNumbersButton;
    }

    partial class ThisRibbonCollection
    {
        internal SerialNumberRibbon Ribbon1
        {
            get { return this.GetRibbon<SerialNumberRibbon>(); }
        }
    }
}
