namespace Excel_Plugin
{
    partial class RibbonCustom : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonCustom()
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
            this.Lbl_welcome = this.Factory.CreateRibbonLabel();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.Btn_TaskPane = this.Factory.CreateRibbonButton();
            this.Btn_WinForm = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.editBox3 = this.Factory.CreateRibbonEditBox();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.checkBox1 = this.Factory.CreateRibbonCheckBox();
            this.checkBox2 = this.Factory.CreateRibbonCheckBox();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            this.tab1.Visible = false;
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Groups.Add(this.group4);
            this.tab2.Groups.Add(this.group5);
            this.tab2.Groups.Add(this.group6);
            this.tab2.Label = "VSTO";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.Lbl_welcome);
            this.group2.Label = "Version Info";
            this.group2.Name = "group2";
            // 
            // Lbl_welcome
            // 
            this.Lbl_welcome.Label = "Lbl_welcome";
            this.Lbl_welcome.Name = "Lbl_welcome";
            // 
            // group3
            // 
            this.group3.Items.Add(this.Btn_TaskPane);
            this.group3.Items.Add(this.Btn_WinForm);
            this.group3.Label = "More control";
            this.group3.Name = "group3";
            // 
            // Btn_TaskPane
            // 
            this.Btn_TaskPane.Label = "Show Task Pane";
            this.Btn_TaskPane.Name = "Btn_TaskPane";
            this.Btn_TaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_TaskPane_Click);
            // 
            // Btn_WinForm
            // 
            this.Btn_WinForm.Label = "Show Win form";
            this.Btn_WinForm.Name = "Btn_WinForm";
            this.Btn_WinForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_WinForm_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.comboBox1);
            this.group4.Items.Add(this.dropDown1);
            this.group4.Label = "Selection";
            this.group4.Name = "group4";
            // 
            // comboBox1
            // 
            this.comboBox1.Label = "comboBox1";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Text = null;
            // 
            // dropDown1
            // 
            this.dropDown1.Label = "dropDown1";
            this.dropDown1.Name = "dropDown1";
            // 
            // group5
            // 
            this.group5.Items.Add(this.editBox1);
            this.group5.Items.Add(this.editBox2);
            this.group5.Items.Add(this.editBox3);
            this.group5.Label = "InputBox";
            this.group5.Name = "group5";
            // 
            // editBox1
            // 
            this.editBox1.Label = "Number 1";
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = null;
            this.editBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_TextChanged);
            // 
            // editBox2
            // 
            this.editBox2.Label = "Number 2";
            this.editBox2.Name = "editBox2";
            this.editBox2.Text = null;
            this.editBox2.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox2_TextChanged);
            // 
            // editBox3
            // 
            this.editBox3.Enabled = false;
            this.editBox3.Label = "Total";
            this.editBox3.Name = "editBox3";
            this.editBox3.Text = null;
            // 
            // group6
            // 
            this.group6.Items.Add(this.checkBox1);
            this.group6.Items.Add(this.checkBox2);
            this.group6.Items.Add(this.toggleButton1);
            this.group6.Label = "CheckBox";
            this.group6.Name = "group6";
            // 
            // checkBox1
            // 
            this.checkBox1.Label = "checkBox1";
            this.checkBox1.Name = "checkBox1";
            // 
            // checkBox2
            // 
            this.checkBox2.Label = "checkBox2";
            this.checkBox2.Name = "checkBox2";
            // 
            // toggleButton1
            // 
            this.toggleButton1.Label = "toggleButton1";
            this.toggleButton1.Name = "toggleButton1";
            // 
            // RibbonCustom
            // 
            this.Name = "RibbonCustom";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonCustom_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel Lbl_welcome;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_TaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_WinForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonCustom RibbonCustom
        {
            get { return this.GetRibbon<RibbonCustom>(); }
        }
    }
}
