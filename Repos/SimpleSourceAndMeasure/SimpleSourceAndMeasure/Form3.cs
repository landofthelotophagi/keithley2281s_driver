using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Ivi.Driver.Interop;
using KeithleyInstruments.Keithley2280.Interop;
using Ivi.DCPwr.Interop;
using Ivi.Dmm.Interop;

namespace SimpleSourceAndMeasure
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        // Constructor (as called from Form 1)
        public Form3(IKeithley2280 driver)
        {
            InitializeComponent();
            keithley = driver;

            // Initial values
            textBox1.Text = keithley.Measurement.Filter.Averaging.Count["1"].ToString();
            comboBox1.Text = keithley.Measurement.Filter.Averaging.Window["1"].ToString();
        }

        // Initialize new driver component
        new IKeithley2280 keithley = new Keithley2280Class();

        // Filter Count TextBox
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            int parsedValue;
            if ((!int.TryParse(textBox1.Text, out parsedValue) && textBox1.Text != ""))
            {
                textBox1.Clear();
                MessageBox.Show("This is a number only field");
                return;
            }
        }

        // Enable Filter Function Button
        private void button1_Click(object sender, EventArgs e)
        {
            if (int.Parse(textBox1.Text) < 2 || int.Parse(textBox1.Text) > 100)
            {
                MessageBox.Show("Filter Function Count Range: 2 - 100 measurements \nValue out of Range \nValue set to minValue = 2 measurements \nFilter Function NOT SET");
                textBox1.Text = "2";
            }
            else
            {
                keithley.Measurement.Filter.Averaging.Count["1"] = int.Parse(textBox1.Text);
                keithley.Measurement.Filter.Averaging.Window["1"] = double.Parse(comboBox1.Text);
                keithley.Measurement.Filter.Averaging.Enabled["1"] = true;
                this.Close();
            }
            
        }
    }
}
