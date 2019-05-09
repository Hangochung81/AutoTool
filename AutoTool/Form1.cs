using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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
        }

        FolderBrowserDialog fbd = new FolderBrowserDialog();
        OpenFileDialog ofd = new OpenFileDialog();
        CollectDataFromHTML cldt = new CollectDataFromHTML();
        //FileDialogCustomPlace ppp = new FileDialogCustomPlace();
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
         
            string[] ColumnName = { "Name", "Date", "Result"};
            string sheetName = platformSheet.Text;
            //cldt.GetSheetName(txtExcelPath.Text);
            cldt.UpdateExcel(sheetName, txtPath.Text, txtExcelPath.Text, dateTimePicker.Value);

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
            
            platformSheet.Items.Clear();
            var listSheetName = cldt.GetSheetName(txtExcelPath.Text);
            
            foreach (var item in listSheetName)
            {
                platformSheet.Items.Add(item);
            }
            platformSheet.Text = listSheetName[0];
        }

        //public bool ValidateFolder()
        //{
        //    if (txtPath.Text == "")
        //    {
        //        txtPath.Focus();
        //        errorProvider1.SetError(txtPath, MessageBox.Show("Please Select your Path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error).ToString());
        //        return false;
        //    }
        //    return true;
        //}


    }
}
