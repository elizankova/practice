using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace It
{
    public partial class mainfrm : Form
    {
        public mainfrm()
        {
            InitializeComponent();
        }

        private void btnnom_Click(object sender, EventArgs e)
        {
            Form1 nomfrm = new Form1();
            nomfrm.ShowDialog();
        }

        private void btnpr_Click(object sender, EventArgs e)
        {
            Form3 prfrm = new Form3();
            prfrm.ShowDialog();
        }

        private void btnrash_Click(object sender, EventArgs e)
        {
            Form4 rash = new Form4();
            rash.ShowDialog();
        }

        private void btnreport_Click(object sender, EventArgs e)
        {
            Form5 rep = new Form5();
            rep.ShowDialog();
        }
    }
}
