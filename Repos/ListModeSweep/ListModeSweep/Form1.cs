using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using KeithleyInstruments.Keithley2280.Interop;

namespace ListModeSweep
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private IKeithley2280 driver = new Keithley2280Class();

        // Initialize and Enable Remote Control Button
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                driver.Initialize(textBox1.Text, true, true, "");
            }
            catch (Exception ex1)
            {
                MessageBox.Show("Not a Valid Instrument Resource Name", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

            if (driver.Initialized == true)
            {
                splitContainer1.Panel2.Enabled = true;
                button3.Enabled = true;
                panel1.Enabled = true;
                comboBox1.SelectedIndex = 0;                
            }
        }

        // Error Query Button
        private void button2_Click(object sender, EventArgs e)
        {
            int code = 0;
            string mess = "";
            driver.Utility.ErrorQuery(ref code, ref mess);
            textBox2.Text = code.ToString();
            textBox3.Text = mess;
        }

        // Instrument Identification Button
        private void button3_Click(object sender, EventArgs e)
        {
            textBox4.Text = driver.Identity.InstrumentManufacturer + " , " + driver.Identity.InstrumentModel;
        }

        // Reset Button
        private void button16_Click(object sender, EventArgs e)
        {
            driver.Utility.Reset();
            comboBox1.SelectedIndex = 0;
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox11.Clear();
        }

        // Close Connection Button
        private void button17_Click(object sender, EventArgs e)
        {
            driver.Close();
            if (driver.Initialized == false)
            {
                splitContainer1.Panel2.Enabled = false;
                button3.Enabled = false;
                panel1.Enabled = false;
            }
        }

        // Run List Sweep Button
        private void button4_Click(object sender, EventArgs e)
        {
            groupBox1.Enabled = true;
            groupBox2.Enabled = true;

            string selectedChannel = comboBox1.Text;

            if (selectedChannel == "Channel 1")
            {
                //Configure the measure function
                driver.Measurement.ConfigureMeasureFunction("1", Keithley2280MeasurementFunctionEnum.Keithley2280MeasurementFunctionConcurrent, Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeAuto, 1.0);
                
                //Configure list
                double[] voltagePoints = textBox5.Text.Split(',').Select(s => double.Parse(s)).ToArray();
                double[] currentPoints = textBox6.Text.Split(',').Select(s => double.Parse(s)).ToArray();
                double[] dwellPoints = textBox7.Text.Split(',').Select(s => double.Parse(s)).ToArray();

                //if((voltagePoints.Count<> == voltagePoints.Count) == voltagePoints.Count)
                driver.Outputs.get_Item("OutputChannel1").List.ConfigureList(1, ref voltagePoints, ref currentPoints, ref dwellPoints);
                
                //Save the list as list 1
                driver.Outputs.get_Item("OutputChannel1").List.SaveList(1);
                
                //Configure the trigger count to match the number of points in the list
                driver.Trigger.set_Count("1", voltagePoints.Count<double>());   //you can specify anyone: voltagePoints or currentPoints or dwellPoints count
                
                //Clear the reading buffer
                driver.Measurement.Buffer.Clear("1");
                
                //Recall list 1 to make it the active list
                driver.Outputs.get_Item("OutputChannel1").List.ConfigureActiveList(1);
                
                //Turn the output ON
                driver.Outputs.get_Item("OutputChannel1").Enabled = true;
                
                //Enable list mode
                driver.Outputs.get_Item("OutputChannel1").List.Enabled = true;
                
                //Set list measurement complete signal to sweep
                driver.Outputs.get_Item("OutputChannel1").List.MComplete = Keithley2280ListMCompleteEnum.Keithley2280ListMCompleteSweep;
                
                //Initiate the trigger model to start the sweep
                driver.Trigger.Initiate("1");
                
                //Wait until the sweep is complete
                driver.Trigger.WaitToComplete(int.Parse(textBox8.Text));
                
                //Fetch the collected data
                double[] sources = { };
                double[] readings = { };
                Keithley2280OutputStateEnum[] modes = { };
                driver.Measurement.Buffer.FetchMultiple("1", ref sources, ref readings, ref modes);

                textBox9.Clear();
                for (int i = 0; i < sources.Length; i++)
                {
                    textBox9.AppendText(sources[i].ToString() +'\n');
                    //textBox9.AppendText(Environment.NewLine);
                }

                textBox10.Clear();
                for (int i = 0; i < readings.Length; i++)
                {
                    textBox10.AppendText(readings[i].ToString() + '\n');
                    //textBox10.AppendText(Environment.NewLine);
                }

                textBox11.Clear();
                for (int i = 0; i < modes.Length; i++)
                {
                    textBox11.AppendText(modes[i].ToString() + '\n');
                    //textBox11.AppendText(Environment.NewLine);
                }
               
                //Turn the output OFF
                driver.Outputs.get_Item("OutputChannel1").Enabled = false;
            }
            else//selectedChannel == "Channel 2"
                if (driver.Identity.InstrumentModel == "MODEL 2280S-32-6" || driver.Identity.InstrumentModel == "MODEL 2281S-20-6" )
                {
                    MessageBox.Show("Not a Valid Channel", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
