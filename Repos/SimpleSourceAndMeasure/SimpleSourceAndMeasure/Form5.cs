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
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }

        // Initialize new driver component
        IKeithley2280 driver = new Keithley2280Class();

        // Constructor (as called from Form 1)
        public Form5(IKeithley2280 keithley2280)
        {
            InitializeComponent();
            driver = keithley2280;

            textBox1.Text = driver.Measurement.Math.MFactor["1"].ToString();
            textBox2.Text = driver.Measurement.Math.BOffset["1"].ToString();
            textBox3.Text = driver.Measurement.Math.Units["1"].Substring(1, 1);
            comboBox1.SelectedItem = "Reading";
        }

        // m Factor TextBox
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            double parsedValue;
            if (!double.TryParse(textBox1.Text, out parsedValue) && textBox1.Text != "" && textBox1.Text != "-")
            {
                MessageBox.Show("This is a number only field");
                textBox1.Clear();
                return;
            }
        }

        // b Offset TextBox
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            double parsedValue;
            if (!double.TryParse(textBox2.Text, out parsedValue) && textBox2.Text != "" && textBox2.Text != "-")
            {
                MessageBox.Show("This is a number only field");
                textBox2.Clear();
                return;
            }
        }

        // Units TextBox
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if ( textBox3.Text.Length > 1)
            {
                MessageBox.Show("Only one character is allowed in the Units suffix name");
                textBox3.Text = textBox3.Text.Substring(0,1);
            }
        }

        // Enable Math Function Button
        private void button1_Click(object sender, EventArgs e)
        {
            if ( double.Parse(textBox1.Text) > 1000000 || double.Parse(textBox1.Text) < -1000000)
            {
                MessageBox.Show("m (Gain) Factor Range: [ -1000000, 1000000] \nValue out Of Range \nValue set to m = 1 (unitary gain)");
                textBox1.Text = "1";
            }
             if (double.Parse(textBox2.Text) > 1000000 || double.Parse(textBox2.Text) < -1000000)
            {
                MessageBox.Show("b Offset Range: [ -1000000, 1000000] \nValue out Of Range \nValue set to b = 0 (zero offset)");
                textBox2.Text = "0";
            }
            driver.Measurement.Math.ConfigureMXB("1", double.Parse(textBox1.Text), double.Parse(textBox2.Text));
            if (string.IsNullOrEmpty(textBox3.Text))
            {
                textBox3.Text = "u";
            }
            driver.Measurement.Math.Units["1"] = textBox3.Text;
            driver.Measurement.Math.Enabled["1"] = true;
            //this.Close();
        }

        // Show Graph Button
        private void button2_Click(object sender, EventArgs e)
        {
            double[] source = new double[100];
            double[] reading = new double[100];
            driver.Measurement.Buffer.FetchDataDouble("1", Keithley2280FormatElementTypeEnum.Keithley2280FormatElementTypeSource, ref source);
            driver.Measurement.Buffer.FetchDataDouble("1", Keithley2280FormatElementTypeEnum.Keithley2280FormatElementTypeReading, ref reading);
            Form2 form2 = new Form2(source, reading);
            form2.ShowDialog();
        }

        // Buffer Data Type ComboBox
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == "CALC")
            {
                button2.Enabled = true;
                driver.Measurement.Buffer.Feed["1"] = Keithley2280MeasureBufferFeedEnum.Keithley2280MeasureBufferFeedCalculate;
            }
            else
            {
                button2.Enabled = false;
                driver.Measurement.Buffer.Feed["1"] = Keithley2280MeasureBufferFeedEnum.Keithley2280MeasureBufferFeedSense;
            }
        }
    }
}
