using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Ivi.Driver.Interop;
using KeithleyInstruments.Keithley2280.Interop;
using Ivi.DCPwr.Interop;
using Ivi.Dmm.Interop;

namespace SimpleSourceAndMeasure
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }

        // Initialize new driver component
        IKeithley2280 driver = new Keithley2280Class();

        double timeout = 0;

        public Form6(IKeithley2280 keithley2280)
        {
            InitializeComponent();
            driver = keithley2280;

            // Set Initial Values

            // Clear Data Buffer TextBoxes
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();

            // Set Output to List Mode

            // Timeout Initial Value
            textBox1.Text = "1000";
            timeout = double.Parse(textBox1.Text);

            // Add saved lists to Saved Lists ComboBox
            for (int i = 1; i < 10; i++)
            {
                if ( driver.Outputs.get_Item("OutputChannel1").List.QueryListLength(i) != 0)
                {
                    comboBox1.Items.Add("List " + i);
                }
            }

            // show active list in Saved Lists Combobox
            MessageBox.Show(driver.Outputs.get_Item("OutputChannel1").List.ActiveList.ToString());

            if (driver.Outputs.get_Item("OutputChannel1").List.ActiveList != 0)
            {
                // Select the active List to be shown in the Select Saved List ComboBox
                comboBox1.SelectedItem = "List " + driver.Outputs.get_Item("OutputChannel1").List.ActiveList.ToString();

                // Get VP, CP, DTP of the active List into arrays
                int listLength = driver.Outputs.get_Item("OutputChannel1").List.QueryListLength(driver.Outputs.get_Item("OutputChannel1").List.ActiveList);
                double[] VP = new double[listLength];
                double[] CP = new double[listLength];
                double[] DTP = new double[listLength];
                driver.Outputs.get_Item("OutputChannel1").List.QueryList(driver.Outputs.get_Item("OutputChannel1").List.ActiveList, ref VP, ref CP, ref DTP );

                // Clear TextBoxes before use
                textBox2.Clear();
                textBox3.Clear();
                textBox4.Clear();

                // Fill the TextBoxes on the left side of the panel
                for ( int i = 0; i < listLength-1; i++ )
                {
                    textBox2.AppendText(VP[i].ToString() + ", ");
                    textBox3.AppendText(CP[i].ToString() + ", ");
                    textBox4.AppendText(DTP[i].ToString() + ", ");
                }
                textBox2.AppendText(VP[listLength-1].ToString());
                textBox3.AppendText(CP[listLength-1].ToString());
                textBox4.AppendText(DTP[listLength-1].ToString());
            }
            
            // Measure Complete ComboBox initial values
            if ( driver.Outputs.get_Item("OutputChannel1").List.MComplete == Keithley2280ListMCompleteEnum.Keithley2280ListMCompleteStep)
            {
                comboBox2.SelectedItem = "after each step";
            }
            else { comboBox2.SelectedItem = "after each sweep"; }

            // Start @ 0 CheckBox initial values
            if (driver.Outputs.get_Item("OutputChannel1").List.StartAtZero)
            {
                checkBox1.Checked = true;
            }
            else { checkBox1.Checked = false; }

            // End @ 0 CheckBox initial values
            if (driver.Outputs.get_Item("OutputChannel1").List.EndAtZero)
            {
                checkBox2.Checked = true;
            }
            else { checkBox2.Checked = false; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Configure the measure function
            driver.Measurement.ConfigureMeasureFunction("1", Keithley2280MeasurementFunctionEnum.Keithley2280MeasurementFunctionConcurrent, Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeAuto, 1.0);

            //Configure list
            double[] voltagePoints = textBox2.Text.Split(',').Select(s => double.Parse(s)).ToArray();
            double[] currentPoints = textBox3.Text.Split(',').Select(s => double.Parse(s)).ToArray();
            double[] dwellPoints = textBox4.Text.Split(',').Select(s => double.Parse(s)).ToArray();

            //if((voltagePoints.Count<> == voltagePoints.Count) == voltagePoints.Count)
            driver.Outputs.get_Item("OutputChannel1").List.ConfigureList(1, ref voltagePoints, ref currentPoints, ref dwellPoints);

            //Save the list as list 1
            driver.Outputs.get_Item("OutputChannel1").List.SaveList(1);

            //driver.Outputs.get_Item("OutputChannel1").List

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
            driver.Trigger.WaitToComplete(int.Parse(textBox1.Text));

            //Fetch the collected data
            double[] sources = { };
            double[] readings = { };
            Keithley2280OutputStateEnum[] modes = { };
            driver.Measurement.Buffer.FetchMultiple("1", ref sources, ref readings, ref modes);

            textBox5.Clear();
            for (int i = 0; i < sources.Length; i++)
            {
                textBox5.AppendText(sources[i].ToString() + '\n');
                //textBox9.AppendText(Environment.NewLine);
            }

            textBox6.Clear();
            for (int i = 0; i < readings.Length; i++)
            {
                textBox6.AppendText(readings[i].ToString() + '\n');
                //textBox10.AppendText(Environment.NewLine);
            }

            textBox7.Clear();
            for (int i = 0; i < modes.Length; i++)
            {
                textBox7.AppendText(modes[i].ToString() + '\n');
                //textBox11.AppendText(Environment.NewLine);
            }

            //Turn the output OFF
            //driver.Outputs.get_Item("OutputChannel1").Enabled = false;
        }
        
        private void Form6_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            driver.Outputs.get_Item("OutputChannel1").List.Enabled = false;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Select the active List to be shown in the Select Saved List ComboBox
            driver.Outputs.get_Item("OutputChannel1").List.ConfigureActiveList(int.Parse(Regex.Match(comboBox1.Text, @"\d+").Value));

            // Get VP, CP, DTP of the active List into arrays
            int listLength = driver.Outputs.get_Item("OutputChannel1").List.QueryListLength(int.Parse(Regex.Match(comboBox1.Text, @"\d+").Value));
            double[] VP = new double[listLength];
            double[] CP = new double[listLength];
            double[] DTP = new double[listLength];
            driver.Outputs.get_Item("OutputChannel1").List.QueryList(driver.Outputs.get_Item("OutputChannel1").List.ActiveList, ref VP, ref CP, ref DTP);

            // Clear TextBoxes before use
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();

            // Fill the TextBoxes on the left side of the panel
            for (int i = 0; i < listLength - 1; i++)
            {
                textBox2.AppendText(VP[i].ToString() + ", ");
                textBox3.AppendText(CP[i].ToString() + ", ");
                textBox4.AppendText(DTP[i].ToString() + ", ");
            }
            textBox2.AppendText(VP[listLength - 1].ToString());
            textBox3.AppendText(CP[listLength - 1].ToString());
            textBox4.AppendText(DTP[listLength - 1].ToString());
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem == "after each step")
            {
                driver.Outputs.get_Item("OutputChannel1").List.MComplete = Keithley2280ListMCompleteEnum.Keithley2280ListMCompleteStep;
            }
            else { driver.Outputs.get_Item("OutputChannel1").List.MComplete = Keithley2280ListMCompleteEnum.Keithley2280ListMCompleteSweep; }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                driver.Outputs.get_Item("OutputChannel1").List.StartAtZero = true;
            }
            else { driver.Outputs.get_Item("OutputChannel1").List.StartAtZero = false; }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                driver.Outputs.get_Item("OutputChannel1").List.EndAtZero = true; 
            }
            else { driver.Outputs.get_Item("OutputChannel1").List.EndAtZero = false; }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            double parsedValue;
            if (!double.TryParse(textBox7.Text, out parsedValue) && textBox7.Text != "")
            {
                MessageBox.Show("This is a number only field");
                textBox1.Clear();
                return;
            }
        }
    }
}
