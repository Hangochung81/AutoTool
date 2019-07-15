using AutoTool.Entity;
using AutoTool.ReportManager;
using AutoTool.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AutoTool
{
    public partial class MainForm : Form
    {
        FolderBrowserDialog fbd = new FolderBrowserDialog();
        OpenFileDialog ofd = new OpenFileDialog();
        ReportUtils cldt = new ReportUtils();

        object reportTypeData = new[] {
                new { Text = "Extent Report (.html)", Value = Constant.REPORT_TYPE_EXTENT },
                new { Text = "TestNG Report (.xml)", Value = Constant.REPORT_TYPE_TESTNG },
                new { Text = "Allure Report (.json)", Value = Constant.REPORT_TYPE_ALLURE }
            };

        object subTitleData = new[] {
                new { Text = "Test Start Time", Value = Constant.TEST_START_TIME_INDEX.ToString() },
                new { Text = "Test End Time", Value = Constant.TEST_END_TIME_INDEX.ToString() },
                new { Text = "Test Status", Value = Constant.TEST_STATUS_INDEX.ToString() },
                new { Text = "Test Detail (Error info)", Value = Constant.TEST_DETAIL_INDEX.ToString() }
            };

        public MainForm()
        {
            InitializeComponent();

            // Set tooltip guideline for template info
            CustomToolTip tip = new CustomToolTip();
            tip.SetToolTip(lblSample1, "Sample1");
            tip.SetToolTip(lblSample2, "Sample2");
            tip.SetToolTip(lblSample3, "Sample3");
            tip.SetToolTip(lblSample4, "Sample4");
            tip.SetToolTip(lblSample5, "Sample5");
            tip.SetToolTip(lblSample6, "Sample6");
            tip.SetToolTip(lblSample7, "Sample7");

            tip.SetToolTip(lblDateRowIndex, "Sample8");
            tip.SetToolTip(lblTestCaseColumnName, "Sample9");
            tip.SetToolTip(lblFillableColumnStartName, "Sample10");
            tip.SetToolTip(lblFillableRowStartIndex, "Sample11");
            tip.SetToolTip(lblColumnNumberPerDate, "Sample12");

            lblSample1.Tag = Properties.Resources.sample1;
            lblSample2.Tag = Properties.Resources.sample2;
            lblSample3.Tag = Properties.Resources.sample3;
            lblSample4.Tag = Properties.Resources.sample4;
            lblSample5.Tag = Properties.Resources.sample5;
            lblSample6.Tag = Properties.Resources.sample6;
            lblSample7.Tag = Properties.Resources.sample7;
            lblDateRowIndex.Tag = Properties.Resources.sample1;
            lblTestCaseColumnName.Tag = Properties.Resources.sample2;
            lblFillableColumnStartName.Tag = Properties.Resources.sample3;
            lblFillableRowStartIndex.Tag = Properties.Resources.sample4;
            lblColumnNumberPerDate.Tag = Properties.Resources.sample5;

            // Prepare init data for Report Type combobox
            cbxReportType.DisplayMember = "Text";
            cbxReportType.ValueMember = "Value";
            cbxReportType.DataSource = reportTypeData;
            cbxReportType.SelectedIndex = 0;
            // Prepare init data for Sub Title combobox
            colSubTitle.DisplayMember = "Text";
            colSubTitle.ValueMember = "Value";
            colSubTitle.DataSource = subTitleData;
        }

        /* ---------------------- USER DEFINED METHODS ---------------------- */
        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void ValidateNumbericTextbox(TextBox control)
        {
            if (String.IsNullOrEmpty(control.Text) || !control.Text.All(char.IsDigit))
            {
                errorProvider.SetError(control, "Please enter numberic");
                successProvider.SetError(control, null);
            }
            else
            {
                errorProvider.SetError(control, null);
                successProvider.SetError(control, "OK");
            }
        }

        private void ValidateLetterTextbox(TextBox control)
        {
            if (String.IsNullOrEmpty(control.Text) || !control.Text.All(char.IsLetter))
            {
                errorProvider.SetError(control, "Please enter letter");
                successProvider.SetError(control, null);
            }
            else
            {
                errorProvider.SetError(control, null);
                successProvider.SetError(control, "OK");
            }
        }

        private void ValidateTemplateInfo()
        {
            // Validate template info
            foreach (Control c in panel2.Controls)
            {
                if (errorProvider.GetError(c).Length > 0)
                {
                    throw new Exception("Some inputted value in template are incorrect. Please check template info again.");
                }
            }
        }

        private void LoadTemplateInfo()
        {
            INIFile ini = new INIFile(Directory.GetCurrentDirectory() + "\\Settings.ini");

            // Load Template config from ini file
            txtTestCaseColumnName.Text = ini.ReadValue("Template", "TestCaseColumnName");
            txtFillableColumnStartName.Text = ini.ReadValue("Template", "FillableColumnStartName");
            txtFillableRowStartIndex.Text = ini.ReadValue("Template", "FillableRowStartIndex");
            txtDateRowIndex.Text = ini.ReadValue("Template", "DateRowIndex");
            txtColumnNumberPerDate.Text = ini.ReadValue("Template", "ColumnNumberPerDate");
            string[] subTitleDataList = ini.ReadValue("Template", "SubTitleColumnIndexList").Split(';');

            dgvSubTitle.Rows.Clear();
            if (subTitleDataList.Count() > 0)
            {
                foreach (var item in subTitleDataList)
                {
                    string[] subTitleInfo = item.Split('-');
                    dgvSubTitle.Rows.Add(subTitleInfo[0], subTitleInfo[1]);
                }
            }

            // Validate Template info
            ValidateNumbericTextbox(txtDateRowIndex);
            ValidateLetterTextbox(txtTestCaseColumnName);
            ValidateLetterTextbox(txtFillableColumnStartName);
            ValidateNumbericTextbox(txtFillableRowStartIndex);
            ValidateNumbericTextbox(txtColumnNumberPerDate);
        }

        private void SaveTemplateInfo()
        {
            INIFile ini = new INIFile(Directory.GetCurrentDirectory() + "\\Settings.ini");

            // Save Template config
            ini.WriteValue("Template", "TestCaseColumnName", txtTestCaseColumnName.Text);
            ini.WriteValue("Template", "FillableColumnStartName", txtFillableColumnStartName.Text);
            ini.WriteValue("Template", "FillableRowStartIndex", txtFillableRowStartIndex.Text);
            ini.WriteValue("Template", "DateRowIndex", txtDateRowIndex.Text);
            ini.WriteValue("Template", "ColumnNumberPerDate", txtColumnNumberPerDate.Text);

            string subTitleList = "";

            // Get chosen sub title list from template info
            for (int i = 0; i < dgvSubTitle.Rows.Count - 1; i++)
            {
                if (dgvSubTitle.Rows[i].Cells[0].Value == null)
                {
                    throw new Exception("Type value of some chosen sub title are empty. Please check template info again.");
                }

                if (dgvSubTitle.Rows[i].Cells[1].Value == null)
                {
                    throw new Exception("Index value of some chosen sub title are empty. Please check template info again.");
                }

                int subTitleType = Convert.ToInt32(dgvSubTitle.Rows[i].Cells[0].Value.ToString());
                int subTitleIndex = Convert.ToInt32(dgvSubTitle.Rows[i].Cells[1].Value.ToString());

                if (subTitleIndex <= 0)
                {
                    throw new Exception("Index value of some chosen sub title = 0 (index must start from 1). Please check template info again.");
                }

                subTitleList += String.Format(";{0}-{1}", subTitleType, subTitleIndex);
            }

            if (subTitleList.Length > 0)
            {
                subTitleList = subTitleList.Substring(1);
            }

            ini.WriteValue("Template", "SubTitleColumnIndexList", subTitleList);
        }

        /* ---------------------- EVENT METHODS ---------------------- */
        private void MainForm_Load(object sender, EventArgs e)
        {
            //Load Template config
            LoadTemplateInfo();

            INIFile ini = new INIFile(Directory.GetCurrentDirectory() + "\\Settings.ini");
            //Load Report config
            txtPath.Text = ini.ReadValue("Report", "FolderPath");
            txtExcelPath.Text = ini.ReadValue("Report", "OutputExcelPath");

            //Load Ignore test case
            string ignoreTestCasePath = ini.ReadValue("Report", "IgnoreTestCasePath");
            if (File.Exists(ignoreTestCasePath) || File.Exists(Directory.GetCurrentDirectory() + "\\" + ignoreTestCasePath))
            {
                if (!File.Exists(ignoreTestCasePath))
                {
                    ignoreTestCasePath = Directory.GetCurrentDirectory() + "\\" + ignoreTestCasePath;
                }
                txtIgnoreTestCase.Text = File.ReadAllText(ignoreTestCasePath);
            }

            //Load target test case
            string targetTestCasePath = ini.ReadValue("Report", "TargetTestCasePath");
            if (File.Exists(targetTestCasePath) || File.Exists(Directory.GetCurrentDirectory() + "\\" + targetTestCasePath))
            {
                if (!File.Exists(targetTestCasePath))
                {
                    targetTestCasePath = Directory.GetCurrentDirectory() + "\\" + targetTestCasePath;
                }
                txtTestCase.Text = File.ReadAllText(targetTestCasePath);
            }

            //Load Guideline document
            webBrowser.Navigate(Directory.GetCurrentDirectory() + "\\AutoCollectGuideline.mht");
        }

        private void Open_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = fbd.SelectedPath;
            }
        }

        private void OpenButton2_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtExcelPath.Text = ofd.FileName;
            }
        }

        private void ExecuteBtn_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                // Validate template info
                ValidateTemplateInfo();

                string[] testCase = null;
                List<KeyValuePair<int, int>> subTitleList = new List<KeyValuePair<int, int>>();

                if (radioButtonChoice.Checked)
                {
                    testCase = txtTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                }
                string[] ignoreTestCase = txtIgnoreTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                
                // Get chosen sub title list from template info
                for (int i = 0; i < dgvSubTitle.Rows.Count - 1; i++)
                {
                    if (dgvSubTitle.Rows[i].Cells[0].Value == null)
                    {
                        throw new Exception("Type value of some chosen sub title are empty. Please check template info again.");
                    }

                    if (dgvSubTitle.Rows[i].Cells[1].Value == null)
                    {
                        throw new Exception("Index value of some chosen sub title are empty. Please check template info again.");
                    }

                    int subTitleType = Convert.ToInt32(dgvSubTitle.Rows[i].Cells[0].Value.ToString());
                    int subTitleIndex = Convert.ToInt32(dgvSubTitle.Rows[i].Cells[1].Value.ToString());

                    if (subTitleIndex <= 0)
                    {
                        throw new Exception("Index value of some chosen sub title = 0 (index must start from 1). Please check template info again.");
                    }

                    subTitleList.Add(new KeyValuePair<int, int> ( subTitleType, subTitleIndex ));
                }

                // Create ReportInfo and TemplateInfo from inputted data
                ReportInfo report = new ReportInfo()
                {
                    SheetName = cbxSheet.Text,
                    ResultPath = txtPath.Text,
                    ReportType = cbxReportType.SelectedValue.ToString(),
                    FilterFile = txtFilterFile.Text,
                    TestSuiteName = txtTestSuite.Text,
                    TargetPath = txtExcelPath.Text,
                    ReportDate = dateTimePicker.Value,
                    TestCaseList = testCase,
                    IgnoreTestCaseList = ignoreTestCase
                };

                TemplateInfo template = new TemplateInfo()
                {
                    TestCaseColumnName = txtTestCaseColumnName.Text,
                    FillableColumnStartName = txtFillableColumnStartName.Text,
                    FillableRowStartIndex = Convert.ToInt32(txtFillableRowStartIndex.Text),
                    DateRowIndex = Convert.ToInt32(txtDateRowIndex.Text),
                    ColumnNumberPerDate = Convert.ToInt32(txtColumnNumberPerDate.Text),
                    SubTitleColumnIndexList = subTitleList
                };

                // Execute updating chosen Excel file with collected data from reports
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                cldt.UpdateExcel(report, template);

                stopwatch.Stop();

                MessageBox.Show(String.Format("Completed in {0} seconds", stopwatch.ElapsedMilliseconds / 1000), "Status", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Open Excel after finish
                if (chxOpenFile.Checked)
                {
                    cldt.OpenExcel(report.TargetPath, report.SheetName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Cursor = Cursors.Arrow;
        }

        private void txtExcelPath_TextChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            cbxSheet.Items.Clear();
            var listSheetName = cldt.GetSheetName(txtExcelPath.Text);
            
            foreach (var item in listSheetName)
            {
                cbxSheet.Items.Add(item);
            }

            if (listSheetName.Count > 0)
            {
                // set default sheet name
                cbxSheet.Text = listSheetName[0];
            }
            
            Cursor = Cursors.Arrow;
        }

        private void radioButtonAll_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonAll.Checked)
            {
                txtTestCase.Enabled = false;
            }
            else
            {
                txtTestCase.Enabled = true;
            }
        }

        private void txtTestCase_TextChanged(object sender, EventArgs e)
        {
            string[] testCaseList = txtTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            lblTestCaseNumber.Text = String.Format("( {0} test cases )", testCaseList.Count());
        }

        private void btnLoadTestCase_Click(object sender, EventArgs e)
        {
            // Load test case list from selected file
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                try
                {
                    txtTestCase.Text = File.ReadAllText(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Cursor = Cursors.Arrow;
            }
        }

        private void btnLoadCurrentTestCase_Click(object sender, EventArgs e)
        {
            // Load current test case list from selected file
            Cursor = Cursors.WaitCursor;
            try
            {
                txtTestCase.Text = cldt.GetCurrentTestCase(txtExcelPath.Text, cbxSheet.Text, txtTestCaseColumnName.Text, Convert.ToInt32(txtFillableRowStartIndex.Text));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor = Cursors.Arrow;
        }

        private void btnEndExcel_Click(object sender, EventArgs e)
        {
            var excelProcs = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in excelProcs)
            {
                proc.Kill();
            }
            MessageBox.Show("Excel processes ended.", "Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void cbxReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxReportType.SelectedIndex == 0)
            {
                lblFileExtension.Text = ".html";
                txtFilterFile.Clear();
                lblTestSuite.Visible = false;
                txtTestSuite.Visible = false;
            }
            else if (cbxReportType.SelectedIndex == 2)
            {
                txtFilterFile.Text = "*result";
                lblFileExtension.Text = ".json";
                lblTestSuite.Visible = false;
                txtTestSuite.Visible = false;
            }
            else
            {
                txtFilterFile.Clear();
                lblFileExtension.Text = ".xml";
                lblTestSuite.Visible = true;
                txtTestSuite.Visible = true;
            }
        }

        private void txtDateRowIndex_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtTestCaseColumnName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }

        }

        private void txtFillableColumnStartName_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsLetter(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtFillableRowStartIndex_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void txtColumnNumberPerDate_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser.Document.Body.Style = "zoom:70%";
        }

        private void btnShowCollectedData_Click(object sender, EventArgs e)
        {
            try
            {
                // Create Report info from inputted value on screen
                string[] ignoreTestCase = txtIgnoreTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                ReportInfo report = new ReportInfo()
                {
                    ResultPath = txtPath.Text,
                    ReportType = cbxReportType.SelectedValue.ToString(),
                    TestSuiteName = txtTestSuite.Text,
                    IgnoreTestCaseList = ignoreTestCase,
                    FilterFile = txtFilterFile.Text
                };

                // get data from reports
                var data = cldt.GetCollectedData(report);
                // show collected data into screen
                HistoryForm form = new HistoryForm(data);
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLoadIgnoreTestCase_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;
                txtIgnoreTestCase.Text = File.ReadAllText(ofd.FileName);
                Cursor = Cursors.Arrow;
            }
        }

        private void txtIgnoreTestCase_TextChanged(object sender, EventArgs e)
        {
            string[] testCaseList = txtIgnoreTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            lblIgnoreTestCaseNumber.Text = String.Format("( {0} test cases )", testCaseList.Count());
        }

        private void dgvSubTitle_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvSubTitle.Columns[e.ColumnIndex] is DataGridViewLinkColumn && e.RowIndex >= 0)
            {
                dgvSubTitle.Rows.RemoveAt(e.RowIndex);
            }
        }

        private void dgvSubTitle_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
        }

        private void txtDateRowIndex_Validating(object sender, CancelEventArgs e)
        {
            ValidateNumbericTextbox(txtDateRowIndex);
        }

        private void txtTestCaseColumnName_Validating(object sender, CancelEventArgs e)
        {
            ValidateLetterTextbox(txtTestCaseColumnName);
        }

        private void txtFillableColumnStartName_Validating(object sender, CancelEventArgs e)
        {
            ValidateLetterTextbox(txtFillableColumnStartName);
        }

        private void txtFillableRowStartIndex_Validating(object sender, CancelEventArgs e)
        {
            ValidateNumbericTextbox(txtFillableRowStartIndex);
        }

        private void txtColumnNumberPerDate_Validating(object sender, CancelEventArgs e)
        {
            ValidateNumbericTextbox(txtColumnNumberPerDate);
        }

        private void btnSaveTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate template info
                ValidateTemplateInfo();

                // Save current template info into Settings.ini file
                SaveTemplateInfo();

                MessageBox.Show("Template was saved.", "Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnReloadTemplate_Click(object sender, EventArgs e)
        {
            // Reload template info
            LoadTemplateInfo();
        }
    }

    class CustomToolTip : ToolTip
    {
        public CustomToolTip()
        {
            this.OwnerDraw = true;
            this.Popup += new PopupEventHandler(this.OnPopup);
            this.Draw += new DrawToolTipEventHandler(this.OnDraw);
        }

        private void OnPopup(object sender, PopupEventArgs e) // use this event to set the size of the tool tip
        {
            e.ToolTipSize = new Size(800, 462);
        }

        private void OnDraw(object sender, DrawToolTipEventArgs e) // use this to customzie the tool tip
        {
            Graphics g = e.Graphics;

            // to set the tag for each button or object
            Control parent = e.AssociatedControl;
            Image pelican = parent.Tag as Image;

            //create your own custom brush to fill the background with the image
            TextureBrush b = new TextureBrush(new Bitmap(pelican));// get the image from Tag

            g.FillRectangle(b, e.Bounds);
            b.Dispose();
        }
    }
}
