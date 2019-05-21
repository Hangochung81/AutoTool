using AutoTool.Utility;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AutoTool
{
    public partial class HistoryForm : Form
    {
        SaveFileDialog sfd = new SaveFileDialog();

        public HistoryForm(List<KeyValuePair<string, string[]>> data)
        {
            InitializeComponent();

            for (int index = 0; index < data.Count; index++)
            {
                dataGridView.Rows.Add(index+1, data.ElementAt(index).Key, data.ElementAt(index).Value[0], data.ElementAt(index).Value[1], data.ElementAt(index).Value[2], data.ElementAt(index).Value[3]);
            }

            // Prepare Save File Dialog Setting
            sfd.Title = "Export Data";
            sfd.DefaultExt = "csv";
            sfd.Filter = "CSV files (*.csv)|*csv";
        }

        private void btnExportData_Click(object sender, EventArgs e)
        {
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                var csv = new StringBuilder();
                csv.AppendLine("Test Name,Start Time,End Time,Status,Detail");
                for (int index = 0; index < dataGridView.RowCount; index++)
                {
                    var newLine = string.Format("{0},{1},{2},{3},{4}", 
                        Helper.AddEscapeSequenceInCsvField(dataGridView.Rows[index].Cells[1].Value.ToString()), 
                        Helper.AddEscapeSequenceInCsvField(dataGridView.Rows[index].Cells[2].Value.ToString()),
                        Helper.AddEscapeSequenceInCsvField(dataGridView.Rows[index].Cells[3].Value.ToString()),
                        Helper.AddEscapeSequenceInCsvField(dataGridView.Rows[index].Cells[4].Value.ToString()),
                        Helper.AddEscapeSequenceInCsvField(dataGridView.Rows[index].Cells[5].Value.ToString()));
                    csv.AppendLine(newLine);
                }

                File.WriteAllText(sfd.FileName, csv.ToString());
            }
        }
    }
}
