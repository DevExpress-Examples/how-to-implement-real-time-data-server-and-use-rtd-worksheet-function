using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestRTDClient
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            spreadsheetControl1.LoadDocument("Portfolio.xlsx");
        }
    }
}
