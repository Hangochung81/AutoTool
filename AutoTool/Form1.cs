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
            CollectDataFromHTML cldt = new CollectDataFromHTML();
            string[] ColumnName = { "Name", "Date", "Result"};
            cldt.UpdateExcel("Sheet1", txtPath.Text, txtExcelPath.Text, ColumnName);
        }
    }
}
