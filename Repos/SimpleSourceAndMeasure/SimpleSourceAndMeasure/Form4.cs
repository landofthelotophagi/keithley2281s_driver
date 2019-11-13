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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        // Initialize new driver component
        IKeithley2280 driver = new Keithley2280Class();

        // Constructor (as called from Form 1)
        public Form4(IKeithley2280 keithley2280)
        {
            InitializeComponent();
            driver = keithley2280;

            // Set Top Label Initial Value
            textBox1.Text = driver.Measurement.Relative.Reference["1"].ToString();
            if (driver.Measurement.MeasureFunction["1"] == Keithley2280MeasurementFunctionEnum.Keithley2280MeasurementFunctionVoltage)
            {
                label1.Text = "Relative Value in Volts";
            }
            else
            {
                label1.Text = "Relative Values in Amperes";
            }
        }
        
        // Reference Value TextBox
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

        // Set Reference Value Button
        private void button1_Click(object sender, EventArgs e)
        {
            if (label1.Text.Contains("Volts"))
            {
                if (double.Parse(textBox1.Text) > 20 || double.Parse(textBox1.Text) < -20)
                {
                    MessageBox.Show("Relative Value Range: [ -20 , 20 ] Volts  \nValue out of Range \nValue set to zero");
                    textBox1.Text = "0";
                }
            }
            else
            {
                if (double.Parse(textBox1.Text) > 6.1 || double.Parse(textBox1.Text) < -6.1)
                {
                    MessageBox.Show("Relative Value Range: [ -6.1 , 6.1 ] Amperes \nValue out of Range \nValue set to zero");
                    textBox1.Text = "0";
                }
            }
            driver.Measurement.Relative.Reference["1"] = double.Parse(textBox1.Text);
        }

        // Acquire Reference Value Button
        private void button2_Click(object sender, EventArgs e)
        {
            driver.Measurement.Relative.Acquire("1");
            textBox1.Text = driver.Measurement.Relative.Reference["1"].ToString();
        }

        // Enable Reference Function Button
        private void button3_Click(object sender, EventArgs e)
        {
            driver.Measurement.Relative.Enabled["1"] = true;
            this.Close();
        }
    }
}
