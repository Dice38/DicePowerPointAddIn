namespace DicePowerPointAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.CL_Button1 = this.Factory.CreateRibbonButton();
            this.CL_DropDown1 = this.Factory.CreateRibbonDropDown();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "Dice Tools";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.label1);
            this.group2.Items.Add(this.CL_DropDown1);
            this.group2.Items.Add(this.CL_Button1);
            this.group2.Name = "group2";
            // 
            // label1
            // 
            this.label1.Label = "Change Language";
            this.label1.Name = "label1";
            // 
            // CL_Button1
            // 
            this.CL_Button1.Label = "Change ";
            this.CL_Button1.Name = "CL_Button1";
            this.CL_Button1.SuperTip = "Changes the language in all TextBoxes, Shapes etc.";
            this.CL_Button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // CL_DropDown1
            // 
            this.CL_DropDown1.Label = "dropDown1";
            this.CL_DropDown1.Name = "CL_DropDown1";
            this.CL_DropDown1.ShowLabel = false;
            this.CL_DropDown1.SizeString = "German";
            this.CL_DropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CL_DropDown1_SelectionChanged);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CL_Button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown CL_DropDown1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
