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
using Keithley2281_Handler_Version1;

namespace SimpleSourceAndMeasure
{
    public partial class InitialForm : Form
    {
        public InitialForm()
        {
            InitializeComponent();
            panel1.Enabled = false;
        }

        // Initialize new driver component
        private IKeithley2280 driver = new Keithley2280Class();
        Class1 dll = new Class1("USB0::0x05E6::0x2281::4380228::INSTR");

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dll.chooseOperationMode("Power Supply");
            }
            catch (Exception ex1)
            {
                MessageBox.Show("Not a Valid Instrument Resource Name", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

            if (driver.Initialized == true)
            {
                driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeEntry;
                panel1.Enabled = true;
                textBox2.Text = driver.Identity.InstrumentManufacturer + " , " + driver.Identity.InstrumentModel;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply;
            Form1 form1 = new Form1(driver);
            form1.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatteryTest;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatterySimulator;
        }
    }
}
