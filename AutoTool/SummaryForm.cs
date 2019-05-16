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
    public partial class SummaryForm : Form
    {
        public SummaryForm(List<KeyValuePair<string, string[]>> data)
        {
            InitializeComponent();

            for (int index = 0; index < data.Count; index++)
            {
                dataGridView.Rows.Add(index+1, data.ElementAt(index).Key, data.ElementAt(index).Value[0], data.ElementAt(index).Value[1], data.ElementAt(index).Value[2], data.ElementAt(index).Value[3]);
            }
        }

        private void SummaryForm_Load(object sender, EventArgs e)
        {

        }
    }
}
