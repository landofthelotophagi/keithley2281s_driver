using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SimpleSourceAndMeasure
{
    public partial class Form2 : Form
    {

        public Form2()
        {
            InitializeComponent();
        }

        // Constructor (as called from Form 1)
        public Form2(double[] sform1, double [] rform1)
        {
            InitializeComponent();
            for (int i=0; i<sform1.Length; i++)
            {
                this.chart1.Series["Source"].Points.AddXY(i, sform1[i]);
                this.chart2.Series["Readings"].Points.AddXY(i, rform1[i]);
                this.chart1.Invalidate();
                this.chart2.Invalidate();
            }   
        }
    }
}
