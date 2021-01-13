using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace OracleTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnOracleCheck_Click(object sender, EventArgs e)
        {

            lbl1.Text = "Running";
            
            bool check = true;

            if(check)
            {
                lbl1.Text = "Pass";
            }
            else
            {
                lbl1.Text = "Fail";
            }
        }
    }
}
