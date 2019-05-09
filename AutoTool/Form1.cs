using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoTool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CustomToolTip tip = new CustomToolTip();
            tip.SetToolTip(lblSample1, "Sample1");
            tip.SetToolTip(lblSample2, "Sample2");
            tip.SetToolTip(lblSample3, "Sample3");
            tip.SetToolTip(lblSample4, "Sample4");
            tip.SetToolTip(lblSample5, "Sample5");
            tip.SetToolTip(lblSample6, "Sample6");
            lblSample1.Tag = Properties.Resources.sample1;
            lblSample2.Tag = Properties.Resources.sample2;
            lblSample3.Tag = Properties.Resources.sample3;
            lblSample4.Tag = Properties.Resources.sample4;
            lblSample5.Tag = Properties.Resources.sample5;
            lblSample6.Tag = Properties.Resources.sample6;
        }

        FolderBrowserDialog fbd = new FolderBrowserDialog();
        OpenFileDialog ofd = new OpenFileDialog();
        CollectDataFromReport cldt = new CollectDataFromReport();

        private void Open_Click(object sender, EventArgs e)
        {
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = fbd.SelectedPath;
            }
            //fbd.ShowDialog();
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
            string[] testCase = null;
            if (radioButtonChoice.Checked)
            {
                testCase = txtTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            }

            ReportInfo report = new ReportInfo()
            {
                SheetName = cbxSheet.Text,
                ResultPath = txtPath.Text,
                ReportType = cbxReportType.SelectedValue.ToString(),
                TargetPath = txtExcelPath.Text,
                ReportDate = dateTimePicker.Value,
                TestCaseList = testCase
            };

            TemplateInfo template = new TemplateInfo()
            {
                TestCaseColumnName = txtTestCaseColumnName.Text,
                FillableColumnStartName = txtFillableColumnStartName.Text,
                FillableRowStartIndex = Convert.ToInt32(txtFillableRowStartIndex.Text),
                DateRowIndex = Convert.ToInt32(txtDateRowIndex.Text),
                ColumnNumberPerDate = Convert.ToInt32(txtColumnNumberPerDate.Text),
                StatusColumnIndexPerDate = Convert.ToInt32(txtStatusColumnIndexPerDate.Text),
                DateTimeFormat = txtDateFormat.Text
            };

            try
            {
                cldt.UpdateExcel(report, template);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtPath_Validating(object sender, CancelEventArgs e)
        {
            //if (txtPath.Text == string.Empty)
            //{
            //    errorProvider1.SetError(txtPath, MessageBox.Show("Please Select your Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error).ToString());
            //    e.Cancel = true;
            //}
        }

        private void ExecuteBtn_Validating(object sender, CancelEventArgs e)
        {
            //if (txtPath.Text == "")
            //{
            //    errorProvider1.SetError(txtPath, MessageBox.Show("Please Select your Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error).ToString());
            //    e.Cancel = true;
            //}
        }

        private void txtExcelPath_TextChanged(object sender, EventArgs e)
        {
            
            cbxSheet.Items.Clear();
            var listSheetName = cldt.GetSheetName(txtExcelPath.Text);
            
            foreach (var item in listSheetName)
            {
                cbxSheet.Items.Add(item);
            }
            cbxSheet.Text = listSheetName[0];
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

        private void Form1_Load(object sender, EventArgs e)
        {
            cbxReportType.DisplayMember = "Text";
            cbxReportType.ValueMember = "Value";

            var items = new[] {
                new { Text = "Extent Report (.html)", Value = "html" },
                new { Text = "TestNG Report (.xml)", Value = "xml" }
            };

            cbxReportType.DataSource = items;
            cbxReportType.SelectedIndex = 0;

            INIFile ini = new INIFile(Directory.GetCurrentDirectory() + "\\Settings.ini");
            //Load Report config
            txtPath.Text = ini.ReadValue("Report", "FolderPath");
            txtExcelPath.Text = ini.ReadValue("Report", "OutputExcelPath");
            //Load Template config
            txtTestCaseColumnName.Text = ini.ReadValue("Template", "TestCaseColumnName");
            txtFillableColumnStartName.Text = ini.ReadValue("Template", "FillableColumnStartName");
            txtFillableRowStartIndex.Text = ini.ReadValue("Template", "FillableRowStartIndex");
            txtDateRowIndex.Text = ini.ReadValue("Template", "DateRowIndex");
            txtColumnNumberPerDate.Text = ini.ReadValue("Template", "ColumnNumberPerDate");
            txtStatusColumnIndexPerDate.Text = ini.ReadValue("Template", "StatusColumnIndexPerDate");
            txtDateFormat.Text = ini.ReadValue("Template", "DateTimeFormat");
        }

        private void txtTestCase_TextChanged(object sender, EventArgs e)
        {
            string[] testCaseList = txtTestCase.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            lblTestCaseNumber.Text = String.Format("( {0} test cases )", testCaseList.Count());
        }

        private void btnLoadTestCase_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTestCase.Text = File.ReadAllText(ofd.FileName);
            }
        }

        private void btnLoadCurrentTestCase_Click(object sender, EventArgs e)
        {
            txtTestCase.Text = cldt.GetCurrentTestCase(txtExcelPath.Text, cbxSheet.Text, txtTestCaseColumnName.Text, Convert.ToInt32(txtFillableRowStartIndex.Text));
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

    class INIFile
    {
        public string path { get; private set; }

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public INIFile(string INIPath)
        {
            path = INIPath;
        }

        public string ReadValue(string Section, string Key)
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, this.path);
            return temp.ToString();
        }
    }

}
