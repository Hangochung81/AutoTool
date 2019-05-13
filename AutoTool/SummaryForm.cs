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
        public SummaryForm(Dictionary<string, string[]> data)
        {
            InitializeComponent();

            for (int index = 0; index < data.Keys.Count; index++)
            {
                dataGridView.Rows.Add(index+1, data.Keys.ElementAt(index), data.Values.ElementAt(index)[0], data.Values.ElementAt(index)[1]);
            }
        }

        private void SummaryForm_Load(object sender, EventArgs e)
        {

        }
    }
}
