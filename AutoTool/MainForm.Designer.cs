namespace AutoTool
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.OpenButton = new System.Windows.Forms.Button();
            this.ExecuteBtn = new System.Windows.Forms.Button();
            this.OpenButton2 = new System.Windows.Forms.Button();
            this.cbxSheet = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageReport = new System.Windows.Forms.TabPage();
            this.chxOpenFile = new System.Windows.Forms.CheckBox();
            this.lblFileExtension = new System.Windows.Forms.Label();
            this.txtFilterFile = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnLoadIgnoreTestCase = new System.Windows.Forms.Button();
            this.lblIgnoreTestCaseNumber = new System.Windows.Forms.Label();
            this.txtIgnoreTestCase = new System.Windows.Forms.TextBox();
            this.btnShowCollectedData = new System.Windows.Forms.Button();
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
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.lblSample5 = new System.Windows.Forms.Label();
            this.lblSample7 = new System.Windows.Forms.Label();
            this.lblSample6 = new System.Windows.Forms.Label();
            this.lblColumnNumberPerDate = new System.Windows.Forms.Label();
            this.txtColumnNumberPerDate = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblSample4 = new System.Windows.Forms.Label();
            this.lblSample3 = new System.Windows.Forms.Label();
            this.lblSample2 = new System.Windows.Forms.Label();
            this.lblSample1 = new System.Windows.Forms.Label();
            this.lblTestCaseColumnName = new System.Windows.Forms.Label();
            this.lblFillableRowStartIndex = new System.Windows.Forms.Label();
            this.txtFillableRowStartIndex = new System.Windows.Forms.TextBox();
            this.lblFillableColumnStartName = new System.Windows.Forms.Label();
            this.txtFillableColumnStartName = new System.Windows.Forms.TextBox();
            this.txtTestCaseColumnName = new System.Windows.Forms.TextBox();
            this.lblDateRowIndex = new System.Windows.Forms.Label();
            this.txtDateRowIndex = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dgvSubTitle = new System.Windows.Forms.DataGridView();
            this.colSubTitle = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colSubIndex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDeleteIcon = new System.Windows.Forms.DataGridViewLinkColumn();
            this.label7 = new System.Windows.Forms.Label();
            this.tabPageAbout = new System.Windows.Forms.TabPage();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.tabPageGuideline = new System.Windows.Forms.TabPage();
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.tabControl1.SuspendLayout();
            this.tabPageReport.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPageTemplate.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSubTitle)).BeginInit();
            this.tabPageAbout.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tabPageGuideline.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 41);
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
            this.txtPath.Size = new System.Drawing.Size(408, 20);
            this.txtPath.TabIndex = 2;
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.Location = new System.Drawing.Point(89, 38);
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.Size = new System.Drawing.Size(408, 20);
            this.txtExcelPath.TabIndex = 3;
            this.txtExcelPath.TextChanged += new System.EventHandler(this.txtExcelPath_TextChanged);
            // 
            // OpenButton
            // 
            this.OpenButton.Location = new System.Drawing.Point(503, 8);
            this.OpenButton.Name = "OpenButton";
            this.OpenButton.Size = new System.Drawing.Size(57, 23);
            this.OpenButton.TabIndex = 4;
            this.OpenButton.Text = "Open";
            this.OpenButton.UseVisualStyleBackColor = true;
            this.OpenButton.Click += new System.EventHandler(this.Open_Click);
            // 
            // ExecuteBtn
            // 
            this.ExecuteBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ExecuteBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExecuteBtn.Location = new System.Drawing.Point(264, 531);
            this.ExecuteBtn.Name = "ExecuteBtn";
            this.ExecuteBtn.Size = new System.Drawing.Size(135, 37);
            this.ExecuteBtn.TabIndex = 5;
            this.ExecuteBtn.Text = "Fill Data Into Output File";
            this.ExecuteBtn.UseVisualStyleBackColor = false;
            this.ExecuteBtn.Click += new System.EventHandler(this.ExecuteBtn_Click);
            // 
            // OpenButton2
            // 
            this.OpenButton2.Location = new System.Drawing.Point(503, 36);
            this.OpenButton2.Name = "OpenButton2";
            this.OpenButton2.Size = new System.Drawing.Size(57, 23);
            this.OpenButton2.TabIndex = 6;
            this.OpenButton2.Text = "Open";
            this.OpenButton2.UseVisualStyleBackColor = true;
            this.OpenButton2.Click += new System.EventHandler(this.OpenButton2_Click);
            // 
            // cbxSheet
            // 
            this.cbxSheet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSheet.FormattingEnabled = true;
            this.cbxSheet.Location = new System.Drawing.Point(89, 64);
            this.cbxSheet.Name = "cbxSheet";
            this.cbxSheet.Size = new System.Drawing.Size(216, 21);
            this.cbxSheet.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 68);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Sheet Name";
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker.Location = new System.Drawing.Point(89, 145);
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
            this.tabControl1.Size = new System.Drawing.Size(579, 600);
            this.tabControl1.TabIndex = 10;
            // 
            // tabPageReport
            // 
            this.tabPageReport.Controls.Add(this.ExecuteBtn);
            this.tabPageReport.Controls.Add(this.chxOpenFile);
            this.tabPageReport.Controls.Add(this.lblFileExtension);
            this.tabPageReport.Controls.Add(this.txtFilterFile);
            this.tabPageReport.Controls.Add(this.label9);
            this.tabPageReport.Controls.Add(this.groupBox2);
            this.tabPageReport.Controls.Add(this.btnShowCollectedData);
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
            this.tabPageReport.Location = new System.Drawing.Point(4, 22);
            this.tabPageReport.Name = "tabPageReport";
            this.tabPageReport.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageReport.Size = new System.Drawing.Size(571, 574);
            this.tabPageReport.TabIndex = 0;
            this.tabPageReport.Text = "Report Info";
            this.tabPageReport.UseVisualStyleBackColor = true;
            // 
            // chxOpenFile
            // 
            this.chxOpenFile.AutoSize = true;
            this.chxOpenFile.Location = new System.Drawing.Point(406, 542);
            this.chxOpenFile.Name = "chxOpenFile";
            this.chxOpenFile.Size = new System.Drawing.Size(124, 17);
            this.chxOpenFile.TabIndex = 28;
            this.chxOpenFile.Text = "Open file when finish";
            this.chxOpenFile.UseVisualStyleBackColor = true;
            // 
            // lblFileExtension
            // 
            this.lblFileExtension.AutoSize = true;
            this.lblFileExtension.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileExtension.Location = new System.Drawing.Point(194, 122);
            this.lblFileExtension.Name = "lblFileExtension";
            this.lblFileExtension.Size = new System.Drawing.Size(0, 13);
            this.lblFileExtension.TabIndex = 27;
            // 
            // txtFilterFile
            // 
            this.txtFilterFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilterFile.Location = new System.Drawing.Point(89, 119);
            this.txtFilterFile.Name = "txtFilterFile";
            this.txtFilterFile.Size = new System.Drawing.Size(104, 20);
            this.txtFilterFile.TabIndex = 26;
            this.txtFilterFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(15, 121);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(60, 13);
            this.label9.TabIndex = 25;
            this.label9.Text = "Filter Name";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnLoadIgnoreTestCase);
            this.groupBox2.Controls.Add(this.lblIgnoreTestCaseNumber);
            this.groupBox2.Controls.Add(this.txtIgnoreTestCase);
            this.groupBox2.Location = new System.Drawing.Point(8, 175);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(555, 171);
            this.groupBox2.TabIndex = 23;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Ignore test case list ( test cases will be ignored when collect data - one line f" +
    "or one test case )";
            // 
            // btnLoadIgnoreTestCase
            // 
            this.btnLoadIgnoreTestCase.Location = new System.Drawing.Point(428, 16);
            this.btnLoadIgnoreTestCase.Name = "btnLoadIgnoreTestCase";
            this.btnLoadIgnoreTestCase.Size = new System.Drawing.Size(124, 23);
            this.btnLoadIgnoreTestCase.TabIndex = 18;
            this.btnLoadIgnoreTestCase.Text = "Load test case from file";
            this.btnLoadIgnoreTestCase.UseVisualStyleBackColor = true;
            this.btnLoadIgnoreTestCase.Click += new System.EventHandler(this.btnLoadIgnoreTestCase_Click);
            // 
            // lblIgnoreTestCaseNumber
            // 
            this.lblIgnoreTestCaseNumber.AutoSize = true;
            this.lblIgnoreTestCaseNumber.Location = new System.Drawing.Point(7, 21);
            this.lblIgnoreTestCaseNumber.Name = "lblIgnoreTestCaseNumber";
            this.lblIgnoreTestCaseNumber.Size = new System.Drawing.Size(76, 13);
            this.lblIgnoreTestCaseNumber.TabIndex = 15;
            this.lblIgnoreTestCaseNumber.Text = "( 0 test cases )";
            // 
            // txtIgnoreTestCase
            // 
            this.txtIgnoreTestCase.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtIgnoreTestCase.Location = new System.Drawing.Point(3, 42);
            this.txtIgnoreTestCase.Multiline = true;
            this.txtIgnoreTestCase.Name = "txtIgnoreTestCase";
            this.txtIgnoreTestCase.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtIgnoreTestCase.Size = new System.Drawing.Size(549, 126);
            this.txtIgnoreTestCase.TabIndex = 12;
            this.txtIgnoreTestCase.WordWrap = false;
            this.txtIgnoreTestCase.TextChanged += new System.EventHandler(this.txtIgnoreTestCase_TextChanged);
            // 
            // btnShowCollectedData
            // 
            this.btnShowCollectedData.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnShowCollectedData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnShowCollectedData.Location = new System.Drawing.Point(133, 531);
            this.btnShowCollectedData.Name = "btnShowCollectedData";
            this.btnShowCollectedData.Size = new System.Drawing.Size(124, 37);
            this.btnShowCollectedData.TabIndex = 22;
            this.btnShowCollectedData.Text = "Show Collected Data";
            this.btnShowCollectedData.UseVisualStyleBackColor = false;
            this.btnShowCollectedData.Click += new System.EventHandler(this.btnShowCollectedData_Click);
            // 
            // btnEndExcel
            // 
            this.btnEndExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnEndExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEndExcel.Location = new System.Drawing.Point(8, 531);
            this.btnEndExcel.Name = "btnEndExcel";
            this.btnEndExcel.Size = new System.Drawing.Size(118, 37);
            this.btnEndExcel.TabIndex = 18;
            this.btnEndExcel.Text = "End Excel Processes";
            this.btnEndExcel.UseVisualStyleBackColor = false;
            this.btnEndExcel.Click += new System.EventHandler(this.btnEndExcel_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 95);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 17;
            this.label5.Text = "Report Type";
            // 
            // cbxReportType
            // 
            this.cbxReportType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxReportType.FormattingEnabled = true;
            this.cbxReportType.Location = new System.Drawing.Point(89, 92);
            this.cbxReportType.Name = "cbxReportType";
            this.cbxReportType.Size = new System.Drawing.Size(216, 21);
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
            this.groupBox1.Location = new System.Drawing.Point(8, 354);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(555, 171);
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
            this.btnLoadTestCase.Location = new System.Drawing.Point(304, 16);
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
            this.txtTestCase.Size = new System.Drawing.Size(549, 126);
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
            this.label4.Location = new System.Drawing.Point(15, 148);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Report Date";
            // 
            // tabPageTemplate
            // 
            this.tabPageTemplate.Controls.Add(this.panel2);
            this.tabPageTemplate.Controls.Add(this.label7);
            this.tabPageTemplate.Location = new System.Drawing.Point(4, 22);
            this.tabPageTemplate.Name = "tabPageTemplate";
            this.tabPageTemplate.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageTemplate.Size = new System.Drawing.Size(571, 574);
            this.tabPageTemplate.TabIndex = 1;
            this.tabPageTemplate.Text = "Template Info";
            this.tabPageTemplate.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label11);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Controls.Add(this.lblSample5);
            this.panel2.Controls.Add(this.lblSample7);
            this.panel2.Controls.Add(this.lblSample6);
            this.panel2.Controls.Add(this.lblColumnNumberPerDate);
            this.panel2.Controls.Add(this.txtColumnNumberPerDate);
            this.panel2.Controls.Add(this.panel1);
            this.panel2.Controls.Add(this.lblSample4);
            this.panel2.Controls.Add(this.lblSample3);
            this.panel2.Controls.Add(this.lblSample2);
            this.panel2.Controls.Add(this.lblSample1);
            this.panel2.Controls.Add(this.lblTestCaseColumnName);
            this.panel2.Controls.Add(this.lblFillableRowStartIndex);
            this.panel2.Controls.Add(this.txtFillableRowStartIndex);
            this.panel2.Controls.Add(this.lblFillableColumnStartName);
            this.panel2.Controls.Add(this.txtFillableColumnStartName);
            this.panel2.Controls.Add(this.txtTestCaseColumnName);
            this.panel2.Controls.Add(this.lblDateRowIndex);
            this.panel2.Controls.Add(this.txtDateRowIndex);
            this.panel2.Controls.Add(this.groupBox3);
            this.panel2.Location = new System.Drawing.Point(6, 6);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(559, 562);
            this.panel2.TabIndex = 25;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(109, 350);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(331, 16);
            this.label11.TabIndex = 53;
            this.label11.Text = "( Assign sub titles will be filled data for each date area )";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(189, 371);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(26, 16);
            this.label10.TabIndex = 52;
            this.label10.Text = "Ex:";
            // 
            // lblSample5
            // 
            this.lblSample5.AutoSize = true;
            this.lblSample5.BackColor = System.Drawing.Color.Transparent;
            this.lblSample5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample5.ForeColor = System.Drawing.Color.Blue;
            this.lblSample5.Location = new System.Drawing.Point(62, 150);
            this.lblSample5.Name = "lblSample5";
            this.lblSample5.Size = new System.Drawing.Size(17, 17);
            this.lblSample5.TabIndex = 46;
            this.lblSample5.Text = "?";
            // 
            // lblSample7
            // 
            this.lblSample7.AutoSize = true;
            this.lblSample7.BackColor = System.Drawing.Color.Transparent;
            this.lblSample7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample7.ForeColor = System.Drawing.Color.Blue;
            this.lblSample7.Location = new System.Drawing.Point(270, 371);
            this.lblSample7.Name = "lblSample7";
            this.lblSample7.Size = new System.Drawing.Size(75, 17);
            this.lblSample7.TabIndex = 43;
            this.lblSample7.Text = "Detail info";
            // 
            // lblSample6
            // 
            this.lblSample6.AutoSize = true;
            this.lblSample6.BackColor = System.Drawing.Color.Transparent;
            this.lblSample6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample6.ForeColor = System.Drawing.Color.Blue;
            this.lblSample6.Location = new System.Drawing.Point(218, 371);
            this.lblSample6.Name = "lblSample6";
            this.lblSample6.Size = new System.Drawing.Size(49, 17);
            this.lblSample6.TabIndex = 42;
            this.lblSample6.Text = "Status";
            // 
            // lblColumnNumberPerDate
            // 
            this.lblColumnNumberPerDate.AutoSize = true;
            this.lblColumnNumberPerDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblColumnNumberPerDate.Location = new System.Drawing.Point(87, 151);
            this.lblColumnNumberPerDate.Name = "lblColumnNumberPerDate";
            this.lblColumnNumberPerDate.Size = new System.Drawing.Size(176, 16);
            this.lblColumnNumberPerDate.TabIndex = 44;
            this.lblColumnNumberPerDate.Text = "Column number of date area";
            // 
            // txtColumnNumberPerDate
            // 
            this.txtColumnNumberPerDate.Location = new System.Drawing.Point(339, 149);
            this.txtColumnNumberPerDate.Name = "txtColumnNumberPerDate";
            this.txtColumnNumberPerDate.Size = new System.Drawing.Size(124, 20);
            this.txtColumnNumberPerDate.TabIndex = 45;
            this.txtColumnNumberPerDate.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtColumnNumberPerDate_KeyPress);
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::AutoTool.Properties.Resources.logo;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.ForeColor = System.Drawing.Color.Black;
            this.panel1.Location = new System.Drawing.Point(62, 420);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(426, 111);
            this.panel1.TabIndex = 24;
            // 
            // lblSample4
            // 
            this.lblSample4.AutoSize = true;
            this.lblSample4.BackColor = System.Drawing.Color.Transparent;
            this.lblSample4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSample4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSample4.ForeColor = System.Drawing.Color.Blue;
            this.lblSample4.Location = new System.Drawing.Point(62, 118);
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
            this.lblSample3.Location = new System.Drawing.Point(62, 86);
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
            this.lblSample2.Location = new System.Drawing.Point(62, 55);
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
            this.lblSample1.Location = new System.Drawing.Point(62, 25);
            this.lblSample1.Name = "lblSample1";
            this.lblSample1.Size = new System.Drawing.Size(17, 17);
            this.lblSample1.TabIndex = 38;
            this.lblSample1.Text = "?";
            // 
            // lblTestCaseColumnName
            // 
            this.lblTestCaseColumnName.AutoSize = true;
            this.lblTestCaseColumnName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestCaseColumnName.Location = new System.Drawing.Point(85, 56);
            this.lblTestCaseColumnName.Name = "lblTestCaseColumnName";
            this.lblTestCaseColumnName.Size = new System.Drawing.Size(227, 16);
            this.lblTestCaseColumnName.TabIndex = 33;
            this.lblTestCaseColumnName.Text = "Test cases\'s name column character";
            // 
            // lblFillableRowStartIndex
            // 
            this.lblFillableRowStartIndex.AutoSize = true;
            this.lblFillableRowStartIndex.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFillableRowStartIndex.Location = new System.Drawing.Point(87, 119);
            this.lblFillableRowStartIndex.Name = "lblFillableRowStartIndex";
            this.lblFillableRowStartIndex.Size = new System.Drawing.Size(139, 16);
            this.lblFillableRowStartIndex.TabIndex = 29;
            this.lblFillableRowStartIndex.Text = "Fillable start row index";
            // 
            // txtFillableRowStartIndex
            // 
            this.txtFillableRowStartIndex.Location = new System.Drawing.Point(339, 119);
            this.txtFillableRowStartIndex.Name = "txtFillableRowStartIndex";
            this.txtFillableRowStartIndex.Size = new System.Drawing.Size(124, 20);
            this.txtFillableRowStartIndex.TabIndex = 30;
            this.txtFillableRowStartIndex.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFillableRowStartIndex_KeyPress);
            // 
            // lblFillableColumnStartName
            // 
            this.lblFillableColumnStartName.AutoSize = true;
            this.lblFillableColumnStartName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFillableColumnStartName.Location = new System.Drawing.Point(87, 87);
            this.lblFillableColumnStartName.Name = "lblFillableColumnStartName";
            this.lblFillableColumnStartName.Size = new System.Drawing.Size(185, 16);
            this.lblFillableColumnStartName.TabIndex = 27;
            this.lblFillableColumnStartName.Text = "Fillable start column character";
            // 
            // txtFillableColumnStartName
            // 
            this.txtFillableColumnStartName.Location = new System.Drawing.Point(339, 86);
            this.txtFillableColumnStartName.Name = "txtFillableColumnStartName";
            this.txtFillableColumnStartName.Size = new System.Drawing.Size(124, 20);
            this.txtFillableColumnStartName.TabIndex = 28;
            this.txtFillableColumnStartName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFillableColumnStartName_KeyPress);
            // 
            // txtTestCaseColumnName
            // 
            this.txtTestCaseColumnName.Location = new System.Drawing.Point(339, 55);
            this.txtTestCaseColumnName.Name = "txtTestCaseColumnName";
            this.txtTestCaseColumnName.Size = new System.Drawing.Size(124, 20);
            this.txtTestCaseColumnName.TabIndex = 26;
            this.txtTestCaseColumnName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTestCaseColumnName_KeyPress);
            // 
            // lblDateRowIndex
            // 
            this.lblDateRowIndex.AutoSize = true;
            this.lblDateRowIndex.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDateRowIndex.Location = new System.Drawing.Point(85, 25);
            this.lblDateRowIndex.Name = "lblDateRowIndex";
            this.lblDateRowIndex.Size = new System.Drawing.Size(157, 16);
            this.lblDateRowIndex.TabIndex = 24;
            this.lblDateRowIndex.Text = "Test case date row index";
            // 
            // txtDateRowIndex
            // 
            this.txtDateRowIndex.Location = new System.Drawing.Point(339, 24);
            this.txtDateRowIndex.Name = "txtDateRowIndex";
            this.txtDateRowIndex.Size = new System.Drawing.Size(124, 20);
            this.txtDateRowIndex.TabIndex = 25;
            this.txtDateRowIndex.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDateRowIndex_KeyPress);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dgvSubTitle);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.groupBox3.Location = new System.Drawing.Point(80, 183);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(391, 165);
            this.groupBox3.TabIndex = 51;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Column index of each sub title in date area";
            // 
            // dgvSubTitle
            // 
            this.dgvSubTitle.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dgvSubTitle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSubTitle.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSubTitle,
            this.colSubIndex,
            this.colDeleteIcon});
            this.dgvSubTitle.Location = new System.Drawing.Point(8, 22);
            this.dgvSubTitle.Name = "dgvSubTitle";
            this.dgvSubTitle.RowHeadersVisible = false;
            this.dgvSubTitle.Size = new System.Drawing.Size(375, 134);
            this.dgvSubTitle.TabIndex = 53;
            this.dgvSubTitle.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSubTitle_CellContentClick);
            this.dgvSubTitle.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgvSubTitle_EditingControlShowing);
            // 
            // colSubTitle
            // 
            this.colSubTitle.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.colSubTitle.HeaderText = "Sub Title";
            this.colSubTitle.Name = "colSubTitle";
            this.colSubTitle.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.colSubTitle.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colSubTitle.Width = 200;
            // 
            // colSubIndex
            // 
            this.colSubIndex.HeaderText = "Index";
            this.colSubIndex.Name = "colSubIndex";
            this.colSubIndex.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // colDeleteIcon
            // 
            this.colDeleteIcon.HeaderText = "";
            this.colDeleteIcon.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.colDeleteIcon.Name = "colDeleteIcon";
            this.colDeleteIcon.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.colDeleteIcon.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colDeleteIcon.Text = "Delete";
            this.colDeleteIcon.TrackVisitedState = false;
            this.colDeleteIcon.UseColumnTextForLinkValue = true;
            this.colDeleteIcon.Width = 70;
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
            this.tabPageAbout.Size = new System.Drawing.Size(571, 574);
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
            this.panel3.Size = new System.Drawing.Size(559, 562);
            this.panel3.TabIndex = 26;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(57, 158);
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
            this.panel4.Location = new System.Drawing.Point(62, 420);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(426, 111);
            this.panel4.TabIndex = 24;
            // 
            // tabPageGuideline
            // 
            this.tabPageGuideline.Controls.Add(this.webBrowser);
            this.tabPageGuideline.Location = new System.Drawing.Point(4, 22);
            this.tabPageGuideline.Name = "tabPageGuideline";
            this.tabPageGuideline.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageGuideline.Size = new System.Drawing.Size(571, 574);
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
            this.webBrowser.Size = new System.Drawing.Size(565, 568);
            this.webBrowser.TabIndex = 0;
            this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.webBrowser_DocumentCompleted);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(579, 600);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Collect Report";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPageReport.ResumeLayout(false);
            this.tabPageReport.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPageTemplate.ResumeLayout(false);
            this.tabPageTemplate.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSubTitle)).EndInit();
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
        private System.Windows.Forms.Label lblSample4;
        private System.Windows.Forms.Label lblSample3;
        private System.Windows.Forms.Label lblSample2;
        private System.Windows.Forms.Label lblSample1;
        private System.Windows.Forms.Label lblTestCaseColumnName;
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
        private System.Windows.Forms.TabPage tabPageGuideline;
        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Button btnShowCollectedData;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnLoadIgnoreTestCase;
        private System.Windows.Forms.Label lblIgnoreTestCaseNumber;
        private System.Windows.Forms.TextBox txtIgnoreTestCase;
        private System.Windows.Forms.TextBox txtFilterFile;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblSample5;
        private System.Windows.Forms.Label lblColumnNumberPerDate;
        private System.Windows.Forms.TextBox txtColumnNumberPerDate;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lblSample6;
        private System.Windows.Forms.Label lblSample7;
        private System.Windows.Forms.DataGridView dgvSubTitle;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.DataGridViewComboBoxColumn colSubTitle;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSubIndex;
        private System.Windows.Forms.DataGridViewLinkColumn colDeleteIcon;
        private System.Windows.Forms.Label lblFileExtension;
        private System.Windows.Forms.CheckBox chxOpenFile;
    }
}

