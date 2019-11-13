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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Initialize new driver component
        private IKeithley2280 driver = new Keithley2280Class();

        Keithley2280MeasurementRangeTypeEnum autoOrManual = Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeAuto;
        double val = 0.01;

        public Form1(IKeithley2280 keithley2280)
        {
            InitializeComponent();

            // Assign active driver from InitialForm to driver object
            driver = keithley2280;

            if (driver.Initialized == true)
            {
                splitContainer1.Panel2.Enabled = true;
                button3.Enabled = true;
                panel1.Enabled = true;
                comboBox1.SelectedIndex = 0;
                button5.Enabled = false;
                button7.Enabled = false;

                //driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatteryTest;

                // Reslution Starting Value
                if (driver.Display.get_Resolution("1").ToString().Contains("6") == true)
                {
                    comboBox4.SelectedItem = "6.5";
                }
                else if ((driver.Display.get_Resolution("1").ToString().Contains("5") == true))
                {
                    comboBox4.SelectedItem = "5.5";
                }
                else { comboBox4.SelectedItem = "4.5"; };

                // Measurement and Range Function Starting Value
                if (driver.Measurement.MeasureFunction["1"].ToString().Contains("Concurrent") == true)
                {
                    comboBox2.SelectedItem = "Concurrent (V+I)";
                    comboBox6.SelectedItem = "10 mA";
                    driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeAuto, 0.001);
                }
                else if (driver.Measurement.MeasureFunction["1"].ToString().Contains("Voltage") == true)
                {
                    comboBox2.SelectedItem = "Voltage (V)";
                    comboBox6.SelectedItem = "20 V";
                    driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeManual, 20);
                }
                else
                {
                    comboBox2.SelectedItem = "Current (I)";
                    comboBox6.SelectedItem = "10 mA";
                    driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeAuto, 0.001);
                }

                // Trigger Mode Starting Value
                if (driver.Outputs.get_Item("OutputChannel1").TriggerSource.ToString().Contains("Immediate"))
                {
                    comboBox5.SelectedItem = "Immediate";
                }
                else if (driver.Outputs.get_Item("OutputChannel1").TriggerSource.ToString().Contains("External"))
                {
                    comboBox5.SelectedItem = "External";
                }
                else { comboBox5.SelectedItem = "Manual"; }

                // Protection Limits Starting Values
                driver.Outputs.get_Item("OutputChannel1").VoltageLimit = 20;
                textBox9.Text = "20";
                driver.Outputs.get_Item("OutputChannel1").ConfigureOCP(true, 6.1);
                textBox8.Text = "6,1";
                driver.Outputs.get_Item("OutputChannel1").ConfigureOVP(true, 21);
                textBox7.Text = "21";

                // Filter Function Starting Value
                driver.Measurement.Filter.Averaging.Enabled["1"] = false;

                // Relative Function Starting Value
                driver.Measurement.Relative.Enabled["1"] = false;


                if (driver.Identity.InstrumentModel == "MODEL 2280S-32-6" || driver.Identity.InstrumentModel == "MODEL 2281S-20-6")
                {
                    if (driver.Outputs.get_Item("OutputChannel1").Enabled == false)
                    {
                        MessageBox.Show("Output Channel 1 is not enabled", "CSharp Application", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        checkBox3.Enabled = false;
                        checkBox4.Enabled = false;
                        checkBox5.Enabled = false;
                    }
                    else
                    {
                        checkBox1.Checked = true;
                        checkBox3.Enabled = true;
                        checkBox4.Enabled = true;
                        checkBox5.Enabled = true;
                    }
                }
            }
        }

        

        // Initialize and Enable Remote Control Button
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                driver.Initialize(textBox1.Text, true, false, "");
            }
            catch (Exception ex1)
            {
                MessageBox.Show("Not a Valid Instrument Resource Name", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

            
        }

        // ErrorQuery Button
        private void button2_Click(object sender, EventArgs e)
        {
            int code = 0;
            string mess = "";
            driver.Utility.ErrorQuery(ref code, ref mess);
            textBox2.Text = code.ToString();
            textBox3.Text = mess;
        }

        // Reset Button
        private void button16_Click(object sender, EventArgs e)
        {
            driver.Utility.Reset();
            comboBox1.SelectedIndex = 0;
            textBox5.Clear();
            textBox6.Clear();

            if (driver.Identity.InstrumentModel == "MODEL 2280S-32-6" || driver.Identity.InstrumentModel == "MODEL 2281S-20-6")
            {
                if (driver.Outputs.get_Item("OutputChannel1").Enabled == false)
                {
                    checkBox1.Checked = false;                    
                }
                else
                {
                    checkBox1.Checked = true;                    
                }
            }
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

        // Output Enabled Checkbox
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                driver.Outputs.get_Item("OutputChannel1").Enabled = true;
                button5.Enabled = true;
                button7.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                checkBox5.Enabled = true;
            }
            else
            {
                driver.Outputs.get_Item("OutputChannel1").Enabled = false;
                button5.Enabled = false;
                button7.Enabled = false;
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                checkBox5.Enabled = false;
            }
        }

        // Set Current Button
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                driver.Outputs.get_Item("OutputChannel1").CurrentLimit = Double.Parse(textBox5.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Invalid Double value", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        // Get Current Button
        private void button5_Click(object sender, EventArgs e)
        {
            textBox5.Text = driver.Outputs.get_Item("OutputChannel1").Measure(Keithley2280MeasurementTypeEnum.Keithley2280MeasurementCurrent).ToString();
        }

        // Set Voltage Button
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                driver.Outputs.get_Item("OutputChannel1").VoltageLevel = Double.Parse(textBox6.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Invalid Double value", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        // Get Voltage Button
        private void button7_Click(object sender, EventArgs e)
        {
            textBox6.Text = driver.Outputs.get_Item("OutputChannel1").Measure(Keithley2280MeasurementTypeEnum.Keithley2280MeasurementVoltage).ToString();
        }
       
        // Instrument Identification Button
        private void button3_Click(object sender, EventArgs e)
        {
            textBox4.Text = driver.Identity.InstrumentManufacturer + " , " + driver.Identity.InstrumentModel;
        }

        // Choose List Button
        private void button11_Click(object sender, EventArgs e)
        {
            Form6 form = new Form6(driver);
            form.ShowDialog();
            MessageBox.Show(driver.Outputs.get_Item("OutputChannel1").List.Enabled.ToString());
        }

        // Clear Data Buffer
        private void button10_Click(object sender, EventArgs e)
        {
            driver.Measurement.Buffer.Clear("1");
        }

        // Show Data Sheet Button
        private void button9_Click(object sender, EventArgs e)
        {
            double[] source = new double[100];
            double[] reading = new double[100];
            driver.Measurement.Buffer.FetchDataDouble("1", Keithley2280FormatElementTypeEnum.Keithley2280FormatElementTypeSource, ref source);
            driver.Measurement.Buffer.FetchDataDouble("1", Keithley2280FormatElementTypeEnum.Keithley2280FormatElementTypeReading, ref reading);
            MessageBox.Show(source.Length.ToString());
            Form2 form2 = new Form2(source, reading);
            form2.ShowDialog();
        }

        // Show Graph Button
        private void button8_Click(object sender, EventArgs e)
        {
            driver.Display.CurrentScreen = Keithley2280DisplayScreenEnum.Keithley2280DisplayScreenGraph;
        }

        // Range Button
        private void button13_Click(object sender, EventArgs e)
        {
            string selItem = comboBox6.SelectedItem.ToString();

            switch (selItem)
            {
                case "10 mA":
                    val = 0.01;
                    break;
                case "100 mA":
                    val = 0.1;
                    break;
                case "1 A":
                    val = 1;
                    break;
                case "10 A":
                    val = 10;
                    break;
            }

            if (comboBox2.SelectedItem.ToString() == "Current" || checkBox2.Checked == false)
            {
                autoOrManual = Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeManual;
            }
            driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], autoOrManual, val);
        }

        // Range Checkbox
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                string selItem = comboBox6.SelectedItem.ToString();
                autoOrManual = Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeAuto;

                switch (selItem)
                {
                    case "10 mA":
                        val = 0.01;
                        break;
                    case "100 mA":
                        val = 0.1;
                        break;
                    case "1 A":
                        val = 1;
                        break;
                    case "10 A":
                        val = 10;
                        break;
                }
                driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], autoOrManual, val);
                comboBox6.ResetText();
                comboBox6.Enabled = false;
                button13.Enabled = false;
            }
            else
            {
                comboBox6.Enabled = true;
                button13.Enabled = true;
                autoOrManual = Keithley2280MeasurementRangeTypeEnum.Keithley2280MeasurementRangeTypeManual;
                if (comboBox2.SelectedItem == "Concurrent (V+I)" || comboBox2.SelectedItem == "Current (I)")
                {
                    comboBox6.SelectedItem = "10 A";
                    driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], autoOrManual, 10);
                }
                else
                {
                    comboBox6.SelectedItem = "20 V";
                    driver.Measurement.ConfigureMeasureFunction("1", driver.Measurement.MeasureFunction["1"], autoOrManual, 20);
                }
            }
        }


        // Measure Combobox
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem == "Concurrent (V+I)")
            {
                driver.Measurement.MeasureFunction["1"] = Keithley2280MeasurementFunctionEnum.Keithley2280MeasurementFunctionConcurrent;
            }
            else if (comboBox2.SelectedItem == "Voltage (V)")
            {
                driver.Measurement.MeasureFunction["1"] = Keithley2280MeasurementFunctionEnum.Keithley2280MeasurementFunctionVoltage;
            }
            else { driver.Measurement.MeasureFunction["1"] = Keithley2280MeasurementFunctionEnum.Keithley2280MeasurementFunctionCurrent; }

            comboBox6.Items.Clear();
            if (comboBox2.SelectedItem == "Concurrent (V+I)" || comboBox2.SelectedItem == "Current (I)")
            {
                comboBox6.Items.Add("10 mA");
                comboBox6.Items.Add("100 mA");
                comboBox6.Items.Add("1 A");
                comboBox6.Items.Add("10 A");
                comboBox6.SelectedItem = "10 mA";
                checkBox2.Enabled = true;
                checkBox2.Checked = true;
                button13.Enabled = true;
            }
            else
            {
                comboBox6.Items.Add("20 V");
                comboBox6.SelectedItem = "20 V";
                checkBox2.Checked = false;
                checkBox2.Enabled = false;
                button13.Enabled = false;
            }
        }


        // Resolution Combobox
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedItem == "4.5")
            {
                driver.Display.set_Resolution("1", Keithley2280ResolutionEnum.Keithley2280Resolution4AndHalf);
            }
            else if (comboBox4.SelectedItem == "5.5")
            {
                driver.Display.set_Resolution("1", Keithley2280ResolutionEnum.Keithley2280Resolution5AndHalf);
            }
            else
            {
                driver.Display.set_Resolution("1", Keithley2280ResolutionEnum.Keithley2280Resolution6AndHalf);
            }
        }


        // Protection Limit Textboxes
        // OVP TextBox
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            double parsedValue;
            if (!double.TryParse(textBox7.Text, out parsedValue) && textBox7.Text != "")
            {
                MessageBox.Show("This is a number only field");
                textBox7.Clear();
                return;
            }
        }

        // Set OVP Button
        private void button14_Click(object sender, EventArgs e)
        {
            if (double.Parse(textBox7.Text) < 0.5 || double.Parse(textBox7.Text) > 21)
            {
                MessageBox.Show("OVP Limit: 0.5 - 21 V \nValue out Of Range \nValue set to maxValue = 21 V");
                textBox7.Text = "21";
            }
            driver.Outputs.get_Item("OutputChannel1").ConfigureOVP(true, Double.Parse(textBox7.Text));
        }

        // OCP TextBox
        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            double parsedValue;
            if (!double.TryParse(textBox8.Text, out parsedValue) && textBox8.Text != "")
            {
                textBox8.Clear();
                MessageBox.Show("This is a number only field");
                return;
            }
        }

        // Set OCP Button
        private void button15_Click(object sender, EventArgs e)
        {
            if (double.Parse(textBox8.Text) < 0.1 || double.Parse(textBox8.Text) > 6.1)
            {
                MessageBox.Show("OCP Limit: 0.1 - 6.1 A \nValue out of Range \nValue set to maxValue = 6.1 A");
                textBox8.Text = "6,1";
            }
            driver.Outputs.get_Item("OutputChannel1").ConfigureOCP(true, Double.Parse(textBox8.Text));
        }

        // VMAX TextBox
        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            double parsedValue;
            if (!double.TryParse(textBox9.Text, out parsedValue) && textBox9.Text != "")
            {
                MessageBox.Show("This is a number only field");
                textBox9.Clear();
                return;
            }
        }
        
        // Set VMAX Button
        private void button18_Click(object sender, EventArgs e)
        {
            if (double.Parse(textBox9.Text) < 0 || double.Parse(textBox9.Text) > 20)
            {
                MessageBox.Show("Vmax Limit: 0 - 20 V \nValue out Of Range \nValue set to maxValue = 20 V");
                textBox9.Text = "20";
            }
            driver.Outputs.get_Item("OutputChannel1").VoltageLimit = Double.Parse(textBox9.Text);
        }


        // Reset Protection Limits Button
        private void button19_Click(object sender, EventArgs e)
        {
            driver.Outputs.get_Item("OutputChannel1").VoltageLimit = 20;
            textBox9.Text = "20";
            driver.Outputs.get_Item("OutputChannel1").ConfigureOCP(true, 6.1);
            textBox8.Text = "6,1";
            driver.Outputs.get_Item("OutputChannel1").ConfigureOVP(true, 21);
            textBox7.Text = "21";
        }


        // Calculate CheckList
        // Relative Function CheckBox
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                Form4 form = new Form4(driver);
                form.ShowDialog();
            }
            else
            {
                driver.Measurement.Relative.Enabled["1"] = false;
            }
        }

        // Math Function CheckBox
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                Form5 form = new Form5(driver);
                form.ShowDialog();
                button9.Enabled = false;
            }
            else
            {
                driver.Measurement.Math.Enabled["1"] = false;
                driver.Measurement.Buffer.Feed["1"] = Keithley2280MeasureBufferFeedEnum.Keithley2280MeasureBufferFeedSense;
                button9.Enabled = true;
            }
        }

        // Filter Function CheckBox
        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                Form3 form3 = new Form3(driver);
                form3.ShowDialog();
            }
            else
            {
                driver.Measurement.Filter.Averaging.Enabled["1"] = false;
            }
        }


        // Acquire Combobox
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedItem == "Immediate")
            {
                driver.Outputs.Item["OutputChannel1"].TriggerSource = Keithley2280TriggerSourceEnum.Keithley2280TriggerSourceImmediate;
                button12.Enabled = false;
            }
            else if (comboBox5.SelectedItem == "External")
            {
                driver.Outputs.Item["OutputChannel1"].TriggerSource = Keithley2280TriggerSourceEnum.Keithley2280TriggerSourceExternal;
                button12.Enabled = true;
            }
            else
            {
                driver.Outputs.Item["OutputChannel1"].TriggerSource = Keithley2280TriggerSourceEnum.Keithley2280TriggerSourceManual;
                button12.Enabled = false;
            }
        }

        // Acquire Button
        private void button12_Click(object sender, EventArgs e)
        {
            driver.Trigger.Initiate("1");
            try
            {
                driver.Trigger.SendSoftwareTrigger();
            }
            catch
            {
                driver.Trigger.Abort();
                MessageBox.Show("Something went wrong. Please recconect with Device.", "CSharpApplication", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
    }
}