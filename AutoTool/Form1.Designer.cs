namespace AutoTool
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.OpenButton = new System.Windows.Forms.Button();
            this.ExecuteBtn = new System.Windows.Forms.Button();
            this.OpenButton2 = new System.Windows.Forms.Button();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.cbxSheet = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageReport = new System.Windows.Forms.TabPage();
            this.btnShowCollectedData = new System.Windows.Forms.Button();
            this.chxUseCustom = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cbxDatetimeFormat = new System.Windows.Forms.ComboBox();
            this.btnEndExcel = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.cbxReportType = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnLoadCurrentTestCase = new System.Windows.Forms.Button();
            this.btnLoadTestCase = new System.Windows.Forms.Button();
            this.lblTestCaseNumber = new System.Windows.Forms.Label();
            this.radioButtonAll = new System.Windows.Forms.RadioButton();
            this.txtTestCase = new System.Windows.Forms.TextBox();
            this.radioButtonChoice = new System.Windows.Forms.RadioButton();
            this.label4 = new System.Windows.Forms.Label();
            this.tabPageTemplate = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblSample6 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblSample5 = new System.Windows.Forms.Label();
            this.lblSample4 = new System.Windows.Forms.Label();
            this.lblSample3 = new System.Windows.Forms.Label();
            this.lblSample2 = new System.Windows.Forms.Label();
            this.lblSample1 = new System.Windows.Forms.Label();
            this.lblColumnNumberPerDate = new System.Windows.Forms.Label();
            this.txtColumnNumberPerDate = new System.Windows.Forms.TextBox();
            this.lblTestCaseColumnName = new System.Windows.Forms.Label();
            this.lblStatusColumnIndexPerDate = new System.Windows.Forms.Label();
            this.txtStatusColumnIndexPerDate = new System.Windows.Forms.TextBox();
            this.lblFillableRowStartIndex = new System.Windows.Forms.Label();
            this.txtFillableRowStartIndex = new System.Windows.Forms.TextBox();
            this.lblFillableColumnStartName = new System.Windows.Forms.Label();
            this.txtFillableColumnStartName = new System.Windows.Forms.TextBox();
            this.txtTestCaseColumnName = new System.Windows.Forms.TextBox();
            this.lblDateRowIndex = new System.Windows.Forms.Label();
            this.txtDateRowIndex = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tabPageAbout = new System.Windows.Forms.TabPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.tabPageGuideline = new System.Windows.Forms.TabPage();
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPageReport.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPageTemplate.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabPageAbout.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tabPageGuideline.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 76);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Output Excel";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Result Folder";
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(89, 10);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(393, 20);
            this.txtPath.TabIndex = 2;
            this.txtPath.Validating += new System.ComponentModel.CancelEventHandler(this.txtPath_Validating);
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.Location = new System.Drawing.Point(89, 74);
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.Size = new System.Drawing.Size(393, 20);
            this.txtExcelPath.TabIndex = 3;
            this.txtExcelPath.TextChanged += new System.EventHandler(this.txtExcelPath_TextChanged);
            // 
            // OpenButton
            // 
            this.OpenButton.Location = new System.Drawing.Point(488, 8);
            this.OpenButton.Name = "OpenButton";
            this.OpenButton.Size = new System.Drawing.Size(75, 23);
            this.OpenButton.TabIndex = 4;
            this.OpenButton.Text = "Open";
            this.OpenButton.UseVisualStyleBackColor = true;
            this.OpenButton.Click += new System.EventHandler(this.Open_Click);
            // 
            // ExecuteBtn
            // 
            this.ExecuteBtn.Location = new System.Drawing.Point(439, 396);
            this.ExecuteBtn.Name = "ExecuteBtn";
            this.ExecuteBtn.Size = new System.Drawing.Size(124, 37);
            this.ExecuteBtn.TabIndex = 5;
            this.ExecuteBtn.Text = "Fill Test Case Status";
            this.ExecuteBtn.UseVisualStyleBackColor = true;
            this.ExecuteBtn.Click += new System.EventHandler(this.ExecuteBtn_Click);
            this.ExecuteBtn.Validating += new System.ComponentModel.CancelEventHandler(this.ExecuteBtn_Validating);
            // 
            // OpenButton2
            // 
            this.OpenButton2.Location = new System.Drawing.Point(488, 71);
            this.OpenButton2.Name = "OpenButton2";
            this.OpenButton2.Size = new System.Drawing.Size(75, 23);
            this.OpenButton2.TabIndex = 6;
            this.OpenButton2.Text = "Open";
            this.OpenButton2.UseVisualStyleBackColor = true;
            this.OpenButton2.Click += new System.EventHandler(this.OpenButton2_Click);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // cbxSheet
            // 
            this.cbxSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSheet.FormattingEnabled = true;
            this.cbxSheet.Location = new System.Drawing.Point(89, 105);
            this.cbxSheet.Name = "cbxSheet";
            this.cbxSheet.Size = new System.Drawing.Size(216, 21);
            this.cbxSheet.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 108);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Sheet Name";
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker.Location = new System.Drawing.Point(89, 140);
            this.dateTimePicker.Name = "dateTimePicker";
            this.dateTimePicker.Size = new System.Drawing.Size(216, 20);
            this.dateTimePicker.TabIndex = 9;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageReport);
            this.tabControl1.Controls.Add(this.tabPageTemplate);
            this.tabControl1.Controls.Add(this.tabPageAbout);
            this.tabControl1.Controls.Add(this.tabPageGuideline);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(579, 467);
            this.tabControl1.TabIndex = 10;
            // 
            // tabPageReport
            // 
            this.tabPageReport.Controls.Add(this.btnShowCollectedData);
            this.tabPageReport.Controls.Add(this.chxUseCustom);
            this.tabPageReport.Controls.Add(this.label8);
            this.tabPageReport.Controls.Add(this.cbxDatetimeFormat);
            this.tabPageReport.Controls.Add(this.btnEndExcel);
            this.tabPageReport.Controls.Add(this.label5);
            this.tabPageReport.Controls.Add(this.cbxReportType);
            this.tabPageReport.Controls.Add(this.groupBox1);
            this.tabPageReport.Controls.Add(this.label4);
            this.tabPageReport.Controls.Add(this.label2);
            this.tabPageReport.Controls.Add(this.dateTimePicker);
            this.tabPageReport.Controls.Add(this.label1);
            this.tabPageReport.Controls.Add(this.label3);
            this.tabPageReport.Controls.Add(this.txtPath);
            this.tabPageReport.Controls.Add(this.cbxSheet);
            this.tabPageReport.Controls.Add(this.txtExcelPath);
            this.tabPageReport.Controls.Add(this.OpenButton2);
            this.tabPageReport.Controls.Add(this.OpenButton);
            this.tabPageReport.Controls.Add(this.ExecuteBtn);
            this.tabPageReport.Location = new System.Drawing.Point(4, 22);
            this.tabPageReport.Name = "tabPageReport";
            this.tabPageReport.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageReport.Size = new System.Drawing.Size(571, 441);
            this.tabPageReport.TabIndex = 0;
            this.tabPageReport.Text = "Report Info";
            this.tabPageReport.UseVisualStyleBackColor = true;
            // 
            // btnShowCollectedData
            // 
            this.btnShowCollectedData.Location = new System.Drawing.Point(132, 396);
            this.btnShowCollectedData.Name = "btnShowCollectedData";
            this.btnShowCollectedData.Size = new System.Drawing.Size(124, 37);
            this.btnShowCollectedData.TabIndex = 22;
            this.btnShowCollectedData.Text = "Show Collected Data";
            this.btnShowCollectedData.UseVisualStyleBackColor = true;
            this.btnShowCollectedData.Click += new System.EventHandler(this.btnShowCollectedData_Click);
            // 
            // chxUseCustom
            // 
            this.chxUseCustom.AutoSize = true;
            this.chxUseCustom.Location = new System.Drawing.Point(467, 43);
            this.chxUseCustom.Name = "chxUseCustom";
            this.chxUseCustom.Size = new System.Drawing.Size(96, 17);
            this.chxUseCustom.TabIndex = 21;
            this.chxUseCustom.Text = "Custom Format";
            this.chxUseCustom.UseVisualStyleBackColor = true;
            this.chxUseCustom.CheckedChanged += new System.EventHandler(this.chxUseCustom_CheckedChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(231, 44);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 13);
            this.label8.TabIndex = 20;
            this.label8.Text = "Date Format";
            // 
            // cbxDatetimeFormat
            // 
            this.cbxDatetimeFormat.Enabled = false;
            this.cbxDatetimeFormat.FormattingEnabled = true;
            this.cbxDatetimeFormat.Items.AddRange(new object[] {
            "M/d/yyyy h:mm:ss tt",
            "yyyy-MM-dd\'T\'HH:mm:ss\'Z\'"});
            this.cbxDatetimeFormat.Location = new System.Drawing.Point(297, 41);
            this.cbxDatetimeFormat.Name = "cbxDatetimeFormat";
            this.cbxDatetimeFormat.Size = new System.Drawing.Size(164, 21);
            this.cbxDatetimeFormat.TabIndex = 19;
            // 
            // btnEndExcel
            // 
            this.btnEndExcel.Location = new System.Drawing.Point(8, 396);
            this.btnEndExcel.Name = "btnEndExcel";
            this.btnEndExcel.Size = new System.Drawing.Size(118, 37);
            this.btnEndExcel.TabIndex = 18;
            this.btnEndExcel.Text = "End Excel processes";
            this.btnEndExcel.UseVisualStyleBackColor = true;
            this.btnEndExcel.Click += new System.EventHandler(this.btnEndExcel_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 44);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Report Type";
            // 
            // cbxReportType
            // 
            this.cbxReportType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxReportType.FormattingEnabled = true;
            this.cbxReportType.Location = new System.Drawing.Point(89, 41);
            this.cbxReportType.Name = "cbxReportType";
            this.cbxReportType.Size = new System.Drawing.Size(131, 21);
            this.cbxReportType.TabIndex = 16;
            this.cbxReportType.SelectedIndexChanged += new System.EventHandler(this.cbxReportType_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnLoadCurrentTestCase);
            this.groupBox1.Controls.Add(this.btnLoadTestCase);
            this.groupBox1.Controls.Add(this.lblTestCaseNumber);
            this.groupBox1.Controls.Add(this.radioButtonAll);
            this.groupBox1.Controls.Add(this.txtTestCase);
            this.groupBox1.Controls.Add(this.radioButtonChoice);
            this.groupBox1.Location = new System.Drawing.Point(8, 176);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(555, 218);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Test case list ( fill status for only chosen test cases in list - one line for on" +
    "e test case )";
            // 
            // btnLoadCurrentTestCase
            // 
            this.btnLoadCurrentTestCase.Location = new System.Drawing.Point(431, 16);
            this.btnLoadCurrentTestCase.Name = "btnLoadCurrentTestCase";
            this.btnLoadCurrentTestCase.Size = new System.Drawing.Size(121, 23);
            this.btnLoadCurrentTestCase.TabIndex = 19;
            this.btnLoadCurrentTestCase.Text = "Load current test case";
            this.btnLoadCurrentTestCase.UseVisualStyleBackColor = true;
            this.btnLoadCurrentTestCase.Click += new System.EventHandler(this.btnLoadCurrentTestCase_Click);
            // 
            // btnLoadTestCase
            // 
            this.btnLoadTestCase.Location = new System.Drawing.Point(297, 16);
            this.btnLoadTestCase.Name = "btnLoadTestCase";
            this.btnLoadTestCase.Size = new System.Drawing.Size(124, 23);
            this.btnLoadTestCase.TabIndex = 18;
            this.btnLoadTestCase.Text = "Load test case from file";
            this.btnLoadTestCase.UseVisualStyleBackColor = true;
            this.btnLoadTestCase.Click += new System.EventHandler(this.btnLoadTestCase_Click);
            // 
            // lblTestCaseNumber
            // 
            this.lblTestCaseNumber.AutoSize = true;
            this.lblTestCaseNumber.Location = new System.Drawing.Point(141, 21);
            this.lblTestCaseNumber.Name = "lblTestCaseNumber";
            this.lblTestCaseNumber.Size = new System.Drawing.Size(76, 13);
            this.lblTestCaseNumber.TabIndex = 15;
            this.lblTestCaseNumber.Text = "( 0 test cases )";
            // 
            // radioButtonAll
            // 
            this.radioButtonAll.AutoSize = true;
            this.radioButtonAll.Checked = true;
            this.radioButtonAll.Location = new System.Drawing.Point(41, 19);
            this.radioButtonAll.Name = "radioButtonAll";
            this.radioButtonAll.Size = new System.Drawing.Size(36, 17);
            this.radioButtonAll.TabIndex = 13;
            this.radioButtonAll.TabStop = true;
            this.radioButtonAll.Text = "All";
            this.radioButtonAll.UseVisualStyleBackColor = true;
            this.radioButtonAll.CheckedChanged += new System.EventHandler(this.radioButtonAll_CheckedChanged);
            // 
            // txtTestCase
            // 
            this.txtTestCase.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtTestCase.Enabled = false;
            this.txtTestCase.Location = new System.Drawing.Point(3, 42);
            this.txtTestCase.Multiline = true;
            this.txtTestCase.Name = "txtTestCase";
            this.txtTestCase.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtTestCase.Size = new System.Drawing.Size(549, 173);
            this.txtTestCase.TabIndex = 12;
            this.txtTestCase.WordWrap = false;
            this.txtTestCase.TextChanged += new System.EventHandler(this.txtTestCase_TextChanged);
            // 
            // radioButtonChoice
            // 
            this.radioButtonChoice.AutoSize = true;
            this.radioButtonChoice.Location = new System.Drawing.Point(85, 19);
            this.radioButtonChoice.Name = "radioButtonChoice";
            this.radioButtonChoice.Size = new System.Drawing.Size(58, 17);
            this.radioButtonChoice.TabIndex = 14;
            this.radioButtonChoice.TabStop = true;
            this.radioButtonChoice.Text = "Choice";
            this.radioButtonChoice.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 144);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Collect Date";
            // 
            // tabPageTemplate
            // 
            this.tabPageTemplate.Controls.Add(this.panel2);
            this.tabPageTemplate.Controls.Add(this.label7);
            this.tabPageTemplate.Location = new System.Drawing.Point(4, 22);
            this.tabPageTemplate.Name = "tabPageTemplate";
            this.tabPageTemplate.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageTemplate.Size = new System.Drawing.Size(571, 441);
            this.tabPageTemplate.TabIndex = 1;
            this.tabPageTemplate.Text = "Template Info";
            this.tabPageTemplate.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.lblSample6);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Controls.Add(this.lblSample5);
            this.panel2.Controls.Add(this.lblSample4);
            this.panel2.Controls.Add(this.lblSample3);
            this.panel2.Controls.Add(this.lblSample2);
            this.panel2.Controls.Add(this.lblSample1);
            this.panel2.Controls.Add(this.lblColumnNumberPerDate);
            this.panel2.Controls.Add(this.txtColumnNumberPerDate);
            this.panel2.Controls.Add(this.lblTestCaseColumnName);
            this.panel2.Controls.Add(this.lblStatusColumnIndexPerDate);
            this.panel2.Controls.Add(this.txtStatusColumnIndexPerDate);
            this.panel2.Controls.Add(this.lblFillableRowStartIndex);
            this.panel2.Controls.Add(this.txtFillableRowStartIndex);
            this.panel2.Controls.Add(this.lblFillableColumnStartName);
            this.panel2.Controls.Add(this.txtFillableColumnStartName);
            this.panel2.Controls.Add(this.txtTestCaseColumnName);
            this.panel2.Controls.Add(this.lblDateRowIndex);
            this.panel2.Controls.Add(this.txtDateRowIndex);
            this.panel2.Location = new System.Drawing.Point(6, 6);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(559, 429);
            this.panel2.TabIndex = 25;
            // 
            // lblSample6
            // 
            this.lblSample6.AutoSize = true;
            this.lblSample6.BackColor = System.Drawing.Color.Transparent;
            this.lblSample6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample6.ForeColor = System.Drawing.Color.Blue;
            this.lblSample6.Location = new System.Drawing.Point(62, 212);
            this.lblSample6.Name = "lblSample6";
            this.lblSample6.Size = new System.Drawing.Size(17, 17);
            this.lblSample6.TabIndex = 43;
            this.lblSample6.Text = "?";
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::AutoTool.Properties.Resources.logo;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.ForeColor = System.Drawing.Color.Black;
            this.panel1.Location = new System.Drawing.Point(123, 305);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(307, 91);
            this.panel1.TabIndex = 24;
            // 
            // lblSample5
            // 
            this.lblSample5.AutoSize = true;
            this.lblSample5.BackColor = System.Drawing.Color.Transparent;
            this.lblSample5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample5.ForeColor = System.Drawing.Color.Blue;
            this.lblSample5.Location = new System.Drawing.Point(62, 176);
            this.lblSample5.Name = "lblSample5";
            this.lblSample5.Size = new System.Drawing.Size(17, 17);
            this.lblSample5.TabIndex = 42;
            this.lblSample5.Text = "?";
            // 
            // lblSample4
            // 
            this.lblSample4.AutoSize = true;
            this.lblSample4.BackColor = System.Drawing.Color.Transparent;
            this.lblSample4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample4.ForeColor = System.Drawing.Color.Blue;
            this.lblSample4.Location = new System.Drawing.Point(62, 137);
            this.lblSample4.Name = "lblSample4";
            this.lblSample4.Size = new System.Drawing.Size(17, 17);
            this.lblSample4.TabIndex = 41;
            this.lblSample4.Text = "?";
            // 
            // lblSample3
            // 
            this.lblSample3.AutoSize = true;
            this.lblSample3.BackColor = System.Drawing.Color.Transparent;
            this.lblSample3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample3.ForeColor = System.Drawing.Color.Blue;
            this.lblSample3.Location = new System.Drawing.Point(62, 100);
            this.lblSample3.Name = "lblSample3";
            this.lblSample3.Size = new System.Drawing.Size(17, 17);
            this.lblSample3.TabIndex = 40;
            this.lblSample3.Text = "?";
            // 
            // lblSample2
            // 
            this.lblSample2.AutoSize = true;
            this.lblSample2.BackColor = System.Drawing.Color.Transparent;
            this.lblSample2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample2.ForeColor = System.Drawing.Color.Blue;
            this.lblSample2.Location = new System.Drawing.Point(62, 63);
            this.lblSample2.Name = "lblSample2";
            this.lblSample2.Size = new System.Drawing.Size(17, 17);
            this.lblSample2.TabIndex = 39;
            this.lblSample2.Text = "?";
            // 
            // lblSample1
            // 
            this.lblSample1.AutoSize = true;
            this.lblSample1.BackColor = System.Drawing.Color.Transparent;
            this.lblSample1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample1.ForeColor = System.Drawing.Color.Blue;
            this.lblSample1.Location = new System.Drawing.Point(62, 27);
            this.lblSample1.Name = "lblSample1";
            this.lblSample1.Size = new System.Drawing.Size(17, 17);
            this.lblSample1.TabIndex = 38;
            this.lblSample1.Text = "?";
            // 
            // lblColumnNumberPerDate
            // 
            this.lblColumnNumberPerDate.AutoSize = true;
            this.lblColumnNumberPerDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblColumnNumberPerDate.Location = new System.Drawing.Point(87, 213);
            this.lblColumnNumberPerDate.Name = "lblColumnNumberPerDate";
            this.lblColumnNumberPerDate.Size = new System.Drawing.Size(176, 16);
            this.lblColumnNumberPerDate.TabIndex = 34;
            this.lblColumnNumberPerDate.Text = "Column number of date area";
            // 
            // txtColumnNumberPerDate
            // 
            this.txtColumnNumberPerDate.Location = new System.Drawing.Point(339, 212);
            this.txtColumnNumberPerDate.Name = "txtColumnNumberPerDate";
            this.txtColumnNumberPerDate.Size = new System.Drawing.Size(110, 20);
            this.txtColumnNumberPerDate.TabIndex = 35;
            this.txtColumnNumberPerDate.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtColumnNumberPerDate_KeyPress);
            // 
            // lblTestCaseColumnName
            // 
            this.lblTestCaseColumnName.AutoSize = true;
            this.lblTestCaseColumnName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestCaseColumnName.Location = new System.Drawing.Point(85, 64);
            this.lblTestCaseColumnName.Name = "lblTestCaseColumnName";
            this.lblTestCaseColumnName.Size = new System.Drawing.Size(227, 16);
            this.lblTestCaseColumnName.TabIndex = 33;
            this.lblTestCaseColumnName.Text = "Test cases\'s name column character";
            // 
            // lblStatusColumnIndexPerDate
            // 
            this.lblStatusColumnIndexPerDate.AutoSize = true;
            this.lblStatusColumnIndexPerDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatusColumnIndexPerDate.Location = new System.Drawing.Point(87, 176);
            this.lblStatusColumnIndexPerDate.Name = "lblStatusColumnIndexPerDate";
            this.lblStatusColumnIndexPerDate.Size = new System.Drawing.Size(233, 16);
            this.lblStatusColumnIndexPerDate.TabIndex = 31;
            this.lblStatusColumnIndexPerDate.Text = "Status column index in each date area";
            // 
            // txtStatusColumnIndexPerDate
            // 
            this.txtStatusColumnIndexPerDate.Location = new System.Drawing.Point(339, 175);
            this.txtStatusColumnIndexPerDate.Name = "txtStatusColumnIndexPerDate";
            this.txtStatusColumnIndexPerDate.Size = new System.Drawing.Size(110, 20);
            this.txtStatusColumnIndexPerDate.TabIndex = 32;
            this.txtStatusColumnIndexPerDate.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtStatusColumnIndexPerDate_KeyPress);
            // 
            // lblFillableRowStartIndex
            // 
            this.lblFillableRowStartIndex.AutoSize = true;
            this.lblFillableRowStartIndex.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFillableRowStartIndex.Location = new System.Drawing.Point(87, 138);
            this.lblFillableRowStartIndex.Name = "lblFillableRowStartIndex";
            this.lblFillableRowStartIndex.Size = new System.Drawing.Size(139, 16);
            this.lblFillableRowStartIndex.TabIndex = 29;
            this.lblFillableRowStartIndex.Text = "Fillable start row index";
            // 
            // txtFillableRowStartIndex
            // 
            this.txtFillableRowStartIndex.Location = new System.Drawing.Point(339, 137);
            this.txtFillableRowStartIndex.Name = "txtFillableRowStartIndex";
            this.txtFillableRowStartIndex.Size = new System.Drawing.Size(110, 20);
            this.txtFillableRowStartIndex.TabIndex = 30;
            this.txtFillableRowStartIndex.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFillableRowStartIndex_KeyPress);
            // 
            // lblFillableColumnStartName
            // 
            this.lblFillableColumnStartName.AutoSize = true;
            this.lblFillableColumnStartName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFillableColumnStartName.Location = new System.Drawing.Point(87, 101);
            this.lblFillableColumnStartName.Name = "lblFillableColumnStartName";
            this.lblFillableColumnStartName.Size = new System.Drawing.Size(185, 16);
            this.lblFillableColumnStartName.TabIndex = 27;
            this.lblFillableColumnStartName.Text = "Fillable start column character";
            // 
            // txtFillableColumnStartName
            // 
            this.txtFillableColumnStartName.Location = new System.Drawing.Point(339, 100);
            this.txtFillableColumnStartName.Name = "txtFillableColumnStartName";
            this.txtFillableColumnStartName.Size = new System.Drawing.Size(110, 20);
            this.txtFillableColumnStartName.TabIndex = 28;
            this.txtFillableColumnStartName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFillableColumnStartName_KeyPress);
            // 
            // txtTestCaseColumnName
            // 
            this.txtTestCaseColumnName.Location = new System.Drawing.Point(339, 63);
            this.txtTestCaseColumnName.Name = "txtTestCaseColumnName";
            this.txtTestCaseColumnName.Size = new System.Drawing.Size(110, 20);
            this.txtTestCaseColumnName.TabIndex = 26;
            this.txtTestCaseColumnName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTestCaseColumnName_KeyPress);
            // 
            // lblDateRowIndex
            // 
            this.lblDateRowIndex.AutoSize = true;
            this.lblDateRowIndex.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDateRowIndex.Location = new System.Drawing.Point(85, 27);
            this.lblDateRowIndex.Name = "lblDateRowIndex";
            this.lblDateRowIndex.Size = new System.Drawing.Size(157, 16);
            this.lblDateRowIndex.TabIndex = 24;
            this.lblDateRowIndex.Text = "Test case date row index";
            // 
            // txtDateRowIndex
            // 
            this.txtDateRowIndex.Location = new System.Drawing.Point(339, 26);
            this.txtDateRowIndex.Name = "txtDateRowIndex";
            this.txtDateRowIndex.Size = new System.Drawing.Size(110, 20);
            this.txtDateRowIndex.TabIndex = 25;
            this.txtDateRowIndex.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDateRowIndex_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(0, 13);
            this.label7.TabIndex = 5;
            // 
            // tabPageAbout
            // 
            this.tabPageAbout.Controls.Add(this.panel3);
            this.tabPageAbout.Location = new System.Drawing.Point(4, 22);
            this.tabPageAbout.Name = "tabPageAbout";
            this.tabPageAbout.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageAbout.Size = new System.Drawing.Size(571, 441);
            this.tabPageAbout.TabIndex = 2;
            this.tabPageAbout.Text = "About";
            this.tabPageAbout.UseVisualStyleBackColor = true;
            // 
            // panel3
            // 
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label6);
            this.panel3.Controls.Add(this.panel4);
            this.panel3.Location = new System.Drawing.Point(6, 6);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(559, 429);
            this.panel3.TabIndex = 26;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(57, 129);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(443, 73);
            this.label6.TabIndex = 25;
            this.label6.Text = "Version: 1.0.0";
            // 
            // panel4
            // 
            this.panel4.BackgroundImage = global::AutoTool.Properties.Resources.logo;
            this.panel4.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.ForeColor = System.Drawing.Color.Black;
            this.panel4.Location = new System.Drawing.Point(123, 305);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(307, 91);
            this.panel4.TabIndex = 24;
            // 
            // tabPageGuideline
            // 
            this.tabPageGuideline.Controls.Add(this.webBrowser);
            this.tabPageGuideline.Location = new System.Drawing.Point(4, 22);
            this.tabPageGuideline.Name = "tabPageGuideline";
            this.tabPageGuideline.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageGuideline.Size = new System.Drawing.Size(571, 441);
            this.tabPageGuideline.TabIndex = 3;
            this.tabPageGuideline.Text = "Guideline";
            this.tabPageGuideline.UseVisualStyleBackColor = true;
            // 
            // webBrowser
            // 
            this.webBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser.Location = new System.Drawing.Point(3, 3);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(565, 435);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_DocumentCompleted);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(579, 467);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Collect Report";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPageReport.ResumeLayout(false);
            this.tabPageReport.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPageTemplate.ResumeLayout(false);
            this.tabPageTemplate.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tabPageAbout.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.tabPageGuideline.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.Button OpenButton;
        private System.Windows.Forms.Button ExecuteBtn;
        private System.Windows.Forms.Button OpenButton2;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbxSheet;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageReport;
        private System.Windows.Forms.TabPage tabPageTemplate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButtonAll;
        private System.Windows.Forms.TextBox txtTestCase;
        private System.Windows.Forms.RadioButton radioButtonChoice;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbxReportType;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblTestCaseNumber;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label lblSample6;
        private System.Windows.Forms.Label lblSample5;
        private System.Windows.Forms.Label lblSample4;
        private System.Windows.Forms.Label lblSample3;
        private System.Windows.Forms.Label lblSample2;
        private System.Windows.Forms.Label lblSample1;
        private System.Windows.Forms.Label lblColumnNumberPerDate;
        private System.Windows.Forms.TextBox txtColumnNumberPerDate;
        private System.Windows.Forms.Label lblTestCaseColumnName;
        private System.Windows.Forms.Label lblStatusColumnIndexPerDate;
        private System.Windows.Forms.TextBox txtStatusColumnIndexPerDate;
        private System.Windows.Forms.Label lblFillableRowStartIndex;
        private System.Windows.Forms.TextBox txtFillableRowStartIndex;
        private System.Windows.Forms.Label lblFillableColumnStartName;
        private System.Windows.Forms.TextBox txtFillableColumnStartName;
        private System.Windows.Forms.TextBox txtTestCaseColumnName;
        private System.Windows.Forms.Label lblDateRowIndex;
        private System.Windows.Forms.TextBox txtDateRowIndex;
        private System.Windows.Forms.Button btnLoadTestCase;
        private System.Windows.Forms.TabPage tabPageAbout;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button btnLoadCurrentTestCase;
        private System.Windows.Forms.Button btnEndExcel;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbxDatetimeFormat;
        private System.Windows.Forms.CheckBox chxUseCustom;
        private System.Windows.Forms.TabPage tabPageGuideline;
        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Button btnShowCollectedData;
    }
}

