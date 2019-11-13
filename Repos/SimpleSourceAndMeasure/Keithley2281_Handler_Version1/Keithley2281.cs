using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.IO.Pipes;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Channels;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Ivi.DCPwr.Interop;
using Ivi.Dmm.Interop;
using Ivi.Driver.Interop;
using KeithleyInstruments.Keithley2280.Interop;
using Timer = System.Threading.Timer;

namespace Keithley2281_Handler_Version1
{
    public class Keithley2281
    {
        // Initialize new driver component
        public IKeithley2280 driver = new Keithley2280();
        private int index = 0;

        // ✔
        // Constructor with arguments
        public Keithley2281(string logicalName)
        {
            try
            {
                driver.Initialize(logicalName, true, false, "");
                DisableOutput();
                Console.WriteLine(driver.Identity.InstrumentManufacturer + " , " + driver.Identity.InstrumentModel);
                return;
            }
            catch (Exception e)
            {
                Console.WriteLine("Not a Valid Instrument Resource Name");
            }
        }

        // Output Control
        // ✔
        private void EnableOutput()
        {
            Keithley2280SystemOperationModeEnum operationMode = driver.System.get_OperationMode("1");
            if (operationMode == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply)
            {
                driver.Outputs.get_Item("1").Enabled = true;
            }
            else { driver.Battery.OutputState["1"] = Keithley2280BatteryOutputStateEnum.Keithley2280BatteryOutputStateOn; }
        }

        // ✔
        private void DisableOutput()
        {
            Keithley2280SystemOperationModeEnum operationMode = driver.System.get_OperationMode("1");
            if (operationMode == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply)
            {
                driver.Outputs.get_Item("1").Enabled = false;
            }
            else { driver.Battery.OutputState["1"] = Keithley2280BatteryOutputStateEnum.Keithley2280BatteryOutputStateOff; }
        }

        // ✔
        // Choose Operation Mode
        public void ChooseOperationMode(string operationMode)
        {
            if (driver.Initialized)
            {
                //channel = chooseChannel();

                driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeEntry; // correct is [channel.ToString()] error -114 when =2

                switch (operationMode)
                {
                    case "Power Supply":
                        powerSupplyMode();
                        break;
                    case "Battery Test":
                        batteryTestMode();
                        break;
                    case "Battery Simulator":
                        BatterySimulatorMode();
                        break;
                    case "Close Connection":
                        CloseConnection();
                        break;
                    default:
                        Console.WriteLine("Wrong String Value");
                        Console.WriteLine("Decide between Power Supply , Battery Test, Battery Simulator and Close Connection");
                        operationMode = Console.ReadLine();
                        ChooseOperationMode(operationMode);
                        break;
                }
            }
            else { CloseConnection(); }
        }

        // ✔
        // Choose Channel
        public int ChooseChannel()  // must include ref value
        {
            int channel = 0;

            Console.WriteLine("Please choose desired channel:");
            Console.WriteLine("1. Channel 1");
            Console.WriteLine("2. Channel 2");

            int x = Convert.ToInt32(Console.ReadLine());
            CheckRange(ref x, 1, 2);

            return channel;
        }

        // ✔
        // Close Connection with Device
        public void CloseConnection()
        {
            if (driver.Initialized)
            {
                driver.Close();
            }
            Console.WriteLine("Session terminated succesfully");
        }

        #region Simple Help Functions - Check Variable Content

        // ✔
        // Check Range Function - change and return int value
        private void CheckRange(ref int val, int margin1, int margin2)
        {
            while (val < margin1 || val > margin2)
            {
                Console.WriteLine("Error. Please try again");
                Console.WriteLine("Value should be between " + margin1.ToString() + " and " + margin2.ToString());
                val = Convert.ToInt32(Console.ReadLine());
            }
        }

        // ✔
        // Check Range Function - simple check double value
        private void CheckRange(ref double val, double margin1, double margin2)
        {
            while (val < margin1 || val > margin2)
            {
                Console.WriteLine("Error. Please try again");
                Console.WriteLine("Value should be between " + margin1.ToString() + " and " + margin2.ToString());
                val = Convert.ToDouble(Console.ReadLine());
            }
        }

        #endregion Simple Help Functions - Check Variable Content

        #region Error Handling

        // ✔
        // Return device errors - NOT YET USED!
        public void DeviceErrors()
        {
            int code = 0;
            string mess = "";
            driver.Utility.ErrorQuery(ref code, ref mess);
            Console.WriteLine("Error code: " + code.ToString());
            Console.WriteLine("Error Message: " + mess);
        }

        // ✔
        // Return connection errors
        public bool ConnectionErrors()
        {
            if (!driver.Initialized)
            {
                Console.WriteLine("Connection with device lost");
                Console.WriteLine("Connection terminated suddenly");
                return true;
            }

            return false;
        }

        #endregion Error Handling

        #region General Purpose Setters, Getters, Settings and Configuration Methods

        // Protection Functions
        // ✔
        // Overvoltage Protection
        public void SetOVP(double val)
        {
            CheckRange(ref val, 0, 21);
            driver.Outputs.get_Item("OutputChannel1").OVPLimit = val;
        }

        // ✔
        public double GetOVP()
        {
            return driver.Outputs.get_Item("OutputChannel1").OVPLimit;
        }

        // ✔
        // Overcurrent Protection
        public void SetOCP(double val)
        {
            CheckRange(ref val, 0.1, 6.1);
            driver.Outputs.get_Item("OutputChannel1").OCPLimit = val;
        }

        // ✔
        public double GetOCP()
        {
            return driver.Outputs.get_Item("OutputChannel1").OCPLimit;
        }

        // ✔
        // Maximum Voltage Protection
        public void SetMAXV(double val)
        {
            CheckRange(ref val, 0, 20);
            driver.Outputs.get_Item("OutputChannel1").VoltageLimit = val;
        }

        // ✔
        public double GetMAXV()
        {
            return driver.Outputs.get_Item("OutputChannel").VoltageLimit;
        }

        // ✔
        // Power Line Frequency
        // READ ONLY
        public Keithley2280PowerLineFrequencyEnum GetPLF()
        {
            return driver.Measurement.PowerLineFrequency;
        }

        // ✔
        // Sample Interval
        // Only for Battery Test and Simulator Mode
        public void SetSampleInterval(double sampInt)
        {
            Keithley2280PowerLineFrequencyEnum plf = GetPLF();
            if (plf == Keithley2280PowerLineFrequencyEnum.Keithley2280PowerLineFrequency50) CheckRange(ref sampInt, 0.00008, 0.48);
            else CheckRange(ref sampInt, 0.0000666667, 0.5);

            if (driver.System.OperationMode["1"] == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatterySimulator)
            {
                string s = sampInt.ToString().Replace(',', '.');
                driver.System.DirectIO.WriteString(":BATT:SIM:SAMP:INT " + s); // Send SCPI Command
            }
            else if (driver.System.OperationMode["1"] == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatteryTest)
            {
                driver.Battery.Test.Measure.SampleInterval["1"] = sampInt;
            }
        }

        // ✔
        public double GetSampleInterval()
        {
            if (driver.System.OperationMode["1"] == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatteryTest)
                return driver.Battery.Test.Measure.SampleInterval["1"];
            else if (driver.System.OperationMode["1"] == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatterySimulator)
            {
                driver.System.DirectIO.WriteString(":BATT:SIM:SAMP:INT?");
                string str = driver.System.DirectIO.ReadString();
                double y = Convert.ToDouble(str.Substring(1, 7).Replace('.',','));
                y = y * Math.Pow(10, Convert.ToDouble(str.Substring(9, 3)));
                return y;
            }

            return 0.125; // Sample Interval mean value in Power Supply Mode
        }

        // ✔
        // Trigger region
        // State :INIT:CONT ON(1) / OFF(0)
        public void SetTriggerState(bool on)
        {
            if (on) { driver.System.DirectIO.WriteString(":INIT:CONT ON"); }
            else { driver.System.DirectIO.WriteString(":INIT:CONT OFF");}
        }

        // ✔
        public string GetTriggerState()
        {
            driver.System.DirectIO.WriteString(":INIT:CONT?");
            return driver.System.DirectIO.ReadString();
        }

        // ✔
        // Arm Count
        public void SetArmCount(int count)
        {
            CheckRange(ref count, 1, 2500);
            driver.Trigger.Arm.Count["1"] = count;
        }

        // ✔
        public int GetArmCount()
        {
            return driver.Trigger.Arm.Count["1"];
        }

        // ✔
        // Trigger Count
        public void SetTriggerCount(int count)
        {
            CheckRange(ref count, 1, 2500);
            driver.Trigger.Count["1"] = count;
        }

        // ✔
        public int GetTriggerCount()
        {
            return driver.Trigger.Count["1"];
        }

        // ✔
        // Trigger and Arm Source (same for both trigger modes)
        public void SetTriggerSource(Keithley2280TriggerSourceEnum source)
        {
            driver.Outputs.get_Item("1").TriggerSource = source;
        }

        // ✔
        public Keithley2280TriggerSourceEnum GeTriggerSource()
        {
            return driver.Outputs.get_Item("1").TriggerSource;
        }

        // ✔
        // Sample Count
        public void SetSampleCount(int sampCount)
        {
            CheckRange(ref sampCount, 1, 2500);
            driver.Measurement.SampleCount["1"] = sampCount;
        }

        // ✔
        public int GetSampleCount()
        {
            return driver.Measurement.SampleCount["1"];
        }

        // ✔
        // List Type
        public void SetTriggerListType(Keithley2280ListMCompleteEnum type)
        {
            if (driver.System.OperationMode["1"] == Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply)
            {
                driver.Outputs.get_Item("1").List.Enabled = true;
                driver.Outputs.get_Item("1").List.MComplete = type;
            }
            else { Console.WriteLine("List is DISABLED in Battery Test and Simulator modes"); driver.Outputs.get_Item("1").List.Enabled = false; } // isws petaei error edw!
        }

        // ✔
        public string GetTriggerListType()
        {
            if (driver.System.OperationMode["1"] ==
                Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply &&
                driver.Outputs.get_Item("1").List.Enabled)
            {
                return driver.Outputs.get_Item("1").List.MComplete.ToString();
            }
            else { return "DISABLED"; }
        }

        #endregion General Purpose Setters, Getters, Settings and Configuration Methods

        #region Power Supply Operation Code

        // All Power Supply Mode Functions and Methods
        public void powerSupplyMode()
        {
            driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply;
            // Thread.Sleep(3000);
            powerSupplyModeClose();
        }

        public void powerSupplyModeClose()
        {
            Console.WriteLine("Choose new operation mode");
            string operationMode = Console.ReadLine();
            ChooseOperationMode(operationMode);
        }

        #endregion Power Supply Operation Code

        #region Battery Test Operation Code

        // ✔
        // All Battery Test Mode Functions and Methods
        public void batteryTestMode()
        {
            driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatteryTest;

            double fullV = 3.2;
            double iLimit = 3.2;

            AHMeasurement(ref fullV, ref iLimit, Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval10Sec);
            batteryTestModeClose();
        }

        // ✔
        public void batteryTestModeClose()
        {
            DisableOutput();
            Console.WriteLine("Choose new operation mode");
            string operationMode = Console.ReadLine();
            ChooseOperationMode(operationMode);
        }

        // ✔
        // pairnei mia instant metrhsh kai thetei meta to output off
        public void ESRMeasurement(ref double evocDelay)
        {
            double esr = 0, voc = 0;

            // Set the ESR and Voc measurements delay parameter
            CheckRange(ref evocDelay, 0, 1200000);
            driver.Battery.Test.Measure.EVOCDelay["1"] = evocDelay;

            driver.Battery.Test.Measure.MeasureEVOC("1", ref esr, ref voc);
        }

        #region AH Measurement and Battery Model Generation

        // ✔
        public void SelectESRInterval(ref Keithley2280BatteryTestMeasureESRIntervalEnum esrInterval)
        {
            Console.WriteLine("Select interval between AH Measurements");
            Console.WriteLine("Available options are:");
            Console.WriteLine("1. 10 seconds");
            Console.WriteLine("2. 30 seconds");
            Console.WriteLine("3. 60 seconds");
            Console.WriteLine("4. 120 seconds");
            Console.WriteLine("5. 10 minutes");
            Console.WriteLine("Please selecto one of the following options by inserting the desired index number");
            int choice = Convert.ToInt32(Console.ReadLine());

            Switch:
            switch (choice)
            {
                case 1:
                    esrInterval = Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval10Sec;
                    break;
                case 2:
                    esrInterval = Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval30Sec;
                    break;
                case 3:
                    esrInterval = Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval60Sec;
                    break;
                case 4:
                    esrInterval = Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval120Sec;
                    break;
                case 5:
                    esrInterval = Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval10Min;
                    break;
                default:
                    Console.WriteLine("Wrong index Value. Please try again");
                    choice = Convert.ToInt32(Console.ReadLine());
                    goto Switch;
            }
        }

        // ✔
        private void FullyDischargeBattery(ref double vLevel, ref double endCond)
        {
            CheckRange(ref vLevel, 0, 21);
            CheckRange(ref endCond, 0.1, 6.1);
            driver.Battery.OutputState["1"] = Keithley2280BatteryOutputStateEnum.Keithley2280BatteryOutputStateOn;
            while (driver.Battery.OutputState["1"] == Keithley2280BatteryOutputStateEnum.Keithley2280BatteryOutputStateOn)
            {
                System.Threading.Thread.Sleep(20000);
            }

            double x = -1;
            double y = -1;
            Keithley2280BatteryTestMeasureESRIntervalEnum z = Keithley2280BatteryTestMeasureESRIntervalEnum.Keithley2280BatteryTestMeasureESRInterval10Min;

            SelectESRInterval(ref z);

            AHMeasurement(ref x, ref y, z);
        }

        // ✔
        private void AHMeasurement(ref double fullV, ref double iLimit, Keithley2280BatteryTestMeasureESRIntervalEnum interval) // ref interval mono ama thelw na kanw kapoio elegxo kai na to allazw endexomenws
        {
            CheckRange(ref fullV, 0, 21);
            driver.Battery.Test.Measure.FullVoltageInAH["1"] = fullV; // Full Voltage

            CheckRange(ref iLimit, 0.1, 6.1);
            driver.Battery.Test.Measure.CurrentLimitInAH["1"] = iLimit; // End Condition

            driver.Battery.Test.Measure.ConfigureAHMeasurement("1", fullV, iLimit, interval);
            driver.Battery.Test.Measure.AHMeasurementState["1"] = Keithley2280BatteryTestAHMeasurementStateEnum.Keithley2280BatteryTestAHMeasurementStateStart; // pote end?
            driver.System.DirectIO.WriteString(":STAT:OPER:INST:ISUM:COND?");
            string response = driver.System.DirectIO.ReadString();
            Console.WriteLine(response);
        }

        // ✔
        public void GenerateBatteryModel(ref double minVOC,ref double maxVOC)
        {
            CheckIndexContent(index);
            CheckRange(ref minVOC, 0, 1); // test
            CheckRange(ref maxVOC, minVOC, 25); // thelw to margin2 > minVOC opwsdhpote
            driver.Battery.Test.Measure.GenerateBatteryModel("1", minVOC, maxVOC, index);

            // mporeis na valeis na rwtaei to xrhsth an thelei na to apothhkeusei sto pc tou. vale kai mia sunarthsh checkYesNo giati einai ntouzntamapnhdes
            SaveBatteryModel(index);
        }

        #endregion AH Measurement and Battery Model Generation

        #endregion Battery Test Operation Code

        #region Battery Simulator Operation Code

        #region Initiate and Close Battery Simulator

        // ✔
        // Initiate Battery Simulator mode
        private void BatterySimulatorMode()
        {
            if (ConnectionErrors()) return; // Check if communication is still intact

            // Change Operation Mode to Battery Simulator - Change Instrument Display Screen
            driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatterySimulator;
        }

        // ✔
        public void BatterySimulatorMode(int i)
        {
            index = i;
            BatterySimulatorMode();
            CheckRange(ref index, 1, 9);
            driver.Battery.Model.RecallByModelIndex("1", index);
        }

        // ✔
        public void BatterySimulatorMode(Keithley2280BatteryBuildInModelEnum builtInModel)
        {
            BatterySimulatorMode();
            driver.Battery.Model.RecallByBuildInModel("1", builtInModel);
        }

        // ✔
        public void BatterySimulatorMode(string filePath)
        {
            BatterySimulatorMode();
            CheckFileContent(ref filePath);
        }

        // ✔
        // Close Battery Simulator mode
        public void BatterySimulatorModeClose()
        {
            DisableOutput();
            driver.System.DirectIO.WriteString(":BATT:TRAC:CLE");
            Console.WriteLine("Write next desired operation mode");
            string operationMode = Console.ReadLine();
            ChooseOperationMode(operationMode);
        }

        #endregion Initiate and Close Battery Simulator

        #region Variable Handling, Setters and Getters

        // ✔
        // Method
        public void setMethod(Keithley2280BatterySimulationModeEnum method)
        {
            driver.Battery.Simulator.SimulationMode["1"] = method;
        }

        // ✔
        public Keithley2280BatterySimulationModeEnum getMethod()
        {
            return driver.Battery.Simulator.SimulationMode["1"];
        }

        // ✔
        // VOC
        public void setVOC(double voc)
        {
            driver.Battery.Simulator.OpenCircuitVoltage["1"] = voc;
        }

        // ✔
        public double getVOC()
        {
            return driver.Battery.Simulator.OpenCircuitVoltage["1"];
        }

        // ✔
        // SOC
        public void setSOC(double soc)
        {
            CheckRange(ref soc, 0, 100);
            driver.Battery.Simulator.StateOfCharge["1"] = soc;
        }

        // ✔
        public double getSOC()
        {
            return driver.Battery.Simulator.StateOfCharge["1"];
        }

        // ✔
        // Full Voltage
        public void setFullV(double fullV)
        {
            driver.Battery.Simulator.OpenCircuitVoltageFull["1"] = fullV;
        }

        // ✔
        public double getFullV()
        {
            return driver.Battery.Simulator.OpenCircuitVoltageFull["1"];
        }

        // ✔
        // Empty Voltage
        public void setEmptyV(double emptyV)
        {
            driver.Battery.Simulator.OpenCircuitVoltageEmpty["1"] = emptyV;
        }

        // ✔
        public double getEmptyV()
        {
            return driver.Battery.Simulator.OpenCircuitVoltageEmpty["1"];
        }

        // ✔
        // I Limit
        public void setILimit(double iLimit)
        {
            driver.Battery.Simulator.CurrentLimit["1"] = iLimit;
        }

        // ✔
        public double GetILimit()
        {
            return driver.Battery.Simulator.CurrentLimit["1"];
        }

        // ✔
        // Capacity Limit
        public void setCapacityLimit(double capLimit)
        {
            driver.Battery.Simulator.CapacityLimit["1"] = capLimit;
        }

        // ✔
        public double getCapacityLimit()
        {
            return driver.Battery.Simulator.CapacityLimit["1"];
        }

        // ✔
        // Resistance Offset
        public void SetResistanceOffset(double resOffset)
        {
            if (driver.Battery.OutputState["1"] == Keithley2280BatteryOutputStateEnum.Keithley2280BatteryOutputStateOff)
            {
                CheckRange(ref resOffset, -100, 100);
                driver.Battery.Simulator.ResistanceOffset["1"] = resOffset;
            }
            else { Console.WriteLine("Simulation is in progress. Resistance Offset value cannot change now!");}

        }

        // ✔
        public double GetResistanceOffset()
        {
            return driver.Battery.Simulator.ResistanceOffset["1"];
        }

        // ✔
        // Status - READ ONLY
        public string GetStatus()
        {
            driver.System.DirectIO.WriteString(":BATT:STAT?");
            return driver.System.DirectIO.ReadString();
        }

        // ✔
        // Get Resistance
        public double getResistance()
        {
            return driver.Battery.Simulator.Resistance["1"];
        }

        // ✔
        // Get Current
        public double getCurrent()
        {
            return driver.Battery.Simulator.Current["1"];
        }

        // ✔
        // Get Capacity
        public double getCapacity()
        {
            return driver.Battery.Simulator.Capacity["1"];
        }

        // ✔
        // Get Terminal Voltage
        public double getTerminalVoltage()
        {
            return driver.Battery.Simulator.TerminalVoltage["1"];
        }

        #endregion Variable Handling, Setters and Getters

        #region Battery Model Handling

        // ✔
        // Chooses whether a Battery Model is in Coarse or Fine Tuning Mode
        private Keithley2280BatteryModelTypeEnum FineOrCoarse(int length)
        {
            if (length == 11) { return Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeCoarse; }
            return Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeFine;
        }

        // ✔
        private int FineOrCoarse(Keithley2280BatteryModelTypeEnum x)
        {
            if (x == Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeCoarse) { return 11; }
            return 101;
        }

        // ✔
        // Returns true if the Battery Model index (arg) is empty 
        private bool CheckIndexContent(int i)
        {
            double[] voc = new double[11];
            double[] res = new double[11];
            double cap = 0;

            driver.Battery.Model.QueryModel("1", Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeCoarse, i, ref voc, ref res, ref cap);

            if (voc[0] == 0){ return true;}

            return false;
        }

        // ✔
        // Returns first available empty index. If all indexes have battery models in store, it returns the value 0
        private int FindEmptyIndex()
        {
            for (int i = 1; i < 10; i++) { if(CheckIndexContent(i)) { return i; }}
            return 0;
        }

        private double[,] SeparateResAndResOffsetArrays(ref double[] res)
        {
            double[,] resOffset = new double[res.Length, 2];

            for (int i = 0; i < res.Length; i++)
            {
                if (res[i] > 10 && res[i] < 20)
                {
                    resOffset[i, 0] = 10;
                    resOffset[i, 1] = 1;
                    res[i] = res[i] - 10;
                }
                else if (res[i] > 20 && res[i] < 110)
                {
                    resOffset[i, 0] = res[i] - 10;
                    resOffset[i, 1] = 2;
                    res[i] = 10;
                }
                else if (res[i] > 110)
                {
                    Console.WriteLine("Value at index " + i.ToString() + " out of instrument resistance bounds");
                    Console.WriteLine("Value set at maximun = 100 Ohm");
                    res[i] = 10;
                    resOffset[i, 0] = 100;
                    resOffset[i, 1] = 3;
                }
            }

            return resOffset;
        }

        public void BatteriesSeries(int n, ref double[] voc, ref double[] res)
        {
            for (int i = 0; i < res.Length; i++)
            {
                voc[i] = n * voc[i];
                res[i] = n * res[i];
            }
        }

        public void BatteriesParallel(int n, ref double[] res, ref double cap)
        {
            for (int i = 0; i < res.Length; i++)
            {
                cap = n * cap;
                res[i] = res[i] / n;
            }
        }

        // ✔
        // Check if desired battery model already exists on device
        private bool CheckIfBatteryModelExists(double[] voc, double[] res, double cap, double[,] resOffset)
        {
            bool flag = true;
            int j = 0;

            double[] modelVoc = new double[voc.Length];
            double[] modelRes = new double[voc.Length];
            double modelCap = 0;
            double[,] modelResOffset = new double[voc.Length, 2];

            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Keithley 2281\Battery Models Repository\ResistanceOffset";
            string extension = ".csv";

            Keithley2280BatteryModelTypeEnum type = FineOrCoarse(voc.Length);

            for (int i = 1; i < 10; i++) // xwris ta built in
            {
                // find resOffset File from Repos
                if (File.Exists(filePath + i.ToString() + extension))
                {
                    List<string> fileLines = new List<string>();
                    StreamReader str = new StreamReader(filePath + i.ToString() + extension);

                    while (!str.EndOfStream)
                    {
                        fileLines.Add(str.ReadLine());
                    }

                    str.Close();

                    string[] lines = fileLines.ToArray();

                    for (j = 0; j < lines.Length; j++)
                    {
                        List<string> separatedValues = new List<string>(lines[j].Split(','));
                        modelResOffset[j, 0] = double.Parse(separatedValues[0].Replace('.', ','));
                        modelResOffset[j, 1] = double.Parse(separatedValues[1]);
                    }
                }

                // Load index model
                driver.Battery.Model.QueryModel("1", type, i, ref modelVoc, ref modelRes, ref modelCap);

                flag = true;

                // Check if same battery model already exists
                if (modelCap == Math.Round(cap, 3))
                {
                    j = 0;
                    while (flag && j<101)
                    {
                        Console.WriteLine(j.ToString() + " , " + Math.Round(voc[j], 4).ToString() + "  ,  " + modelVoc[j].ToString());
                        Console.WriteLine(j.ToString() + " , " + Math.Round(res[j], 4).ToString() + "  ,  " + modelRes[j].ToString());
                        Console.WriteLine(j.ToString() + " , " + resOffset[j,0].ToString() + "  ,  " + modelResOffset[j,0].ToString());
                        if (modelVoc[j] != Math.Round(voc[j], 4)) // || modelRes[j] != Math.Round(res[j], 4) || modelResOffset[j, 0] != resOffset[j, 0])
                        {
                            if(modelRes[j] != Math.Round(res[j], 4)) { }
                            if(modelResOffset[j, 0] != resOffset[j, 0]) { }
                            flag = false;
                            break;
                        }
                        j++;
                    }

                    if (flag)
                    {
                        Console.WriteLine("Battery Model already exists on index " + i.ToString());
                        driver.Battery.Model.RecallByModelIndex("1", i);
                        Console.WriteLine("Battery Model Loaded");
                        return true;
                    }
                }
            }

            return false;
        }

        // ✔
        // Loads one of the built in Battery Models
        public void LoadBuiltInModel()
        {
            Console.WriteLine("The available built-in models are:");
            Console.WriteLine("1. 12 Volts Nickel Metal Hydride");
            Console.WriteLine("2. 1 Point 2 Volts Nickel Cadmium Battery");
            Console.WriteLine("3. 1 Point 2 Volts Nickel Metal Hydride Battery");
            Console.WriteLine("4. 4 Point 2 Volts Lithium Ion Battery");
            Console.WriteLine("5. Lead Acid Battery");
            Console.WriteLine("Please type int he desired model. In order to exit press 0.");

            int batModel = int.Parse(Console.ReadLine());

            switch (batModel)
            {
                case 0:
                    Console.WriteLine("No built-in model was chosen. Exit was successful");
                    break;
                case 1:
                    driver.Battery.Model.RecallByBuildInModel("1", Keithley2280BatteryBuildInModelEnum.Keithley2280BuildInModel12VoltsNickelMetalHydrideBattery);
                    break;
                case 2:
                    driver.Battery.Model.RecallByBuildInModel("1", Keithley2280BatteryBuildInModelEnum.Keithley2280BuildInModel1Point2VoltsNickelCadmiumBattery);
                    break;
                case 3:
                    driver.Battery.Model.RecallByBuildInModel("1", Keithley2280BatteryBuildInModelEnum.Keithley2280BuildInModel1Point2VoltsNickelMetalHydrideBattery);
                    break;
                case 4:
                    driver.Battery.Model.RecallByBuildInModel("1", Keithley2280BatteryBuildInModelEnum.Keithley2280BuildInModel4Point2VoltsLithiumIonBattery);
                    break;
                case 5:
                    driver.Battery.Model.RecallByBuildInModel("1", Keithley2280BatteryBuildInModelEnum.Keithley2280BuildInModelLeadAcidBattery);
                    break;
                default:
                    Console.WriteLine("This Battery Model is not a Keithley 2281S-20-6 built-in model");
                    LoadBuiltInModel();
                    break;
            }
        }

        // ✔
        // Checks whether the content of the desired file can be converted to a model according to the instrument's specification
        private void CheckFileContent(ref string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine("This file does not exist");
                // diaforetikh antimetwpish!!!
                Console.WriteLine("Enter new filepath");
                filePath = Console.ReadLine();
                CheckFileContent(ref filePath);
            }
            else
            {
                StreamReader fileReader = new StreamReader(filePath);
                if (!filePath.EndsWith(".csv") && !filePath.EndsWith(".mdl") && !filePath.EndsWith(".txt"))
                {
                    Console.WriteLine("File type not supported. Please convert to .csv, .mdl or .txt");
                    return;
                    // kapoiou eidous handling!
                }
                else { Console.WriteLine("Welcome to Tijuana."); CreateBatteryModelFromCSVFile(fileReader); }
            }
        } // format check

        // evala assignment to index - na allaksei sto documentation
        // Creates and saves a .csv file with the Battery Model data to the computer
        // NEEDS WORK! // prepei na ftiaksw to kommati me thn onomatlogia twn arxeiwn!
        public void SaveBatteryModel(int x)
        {
            index = x;
            CheckRange(ref index, 1, 9);
            Keithley2280BatteryModelTypeEnum type = FineOrCoarse(index);
            int length = FineOrCoarse(type);
            double[] voc = new double[length];
            double[] res = new double[length];
            double cap = 0;

            // COMMAND CAUSES ERROR
            driver.Battery.Model.QueryModel("1", type, index, ref voc, ref res, ref cap);

            string[] lines = new string[length+3];
            lines[0] = "Battery Model saved at index " + index.ToString() + " of Keithley 2281S-20-6 on " + DateTime.Now.ToString(); 
            lines[1] = "Capacity=" + cap.ToString() + "AH";
            lines[2] = "SOC,Voc,Res";

            switch (type)
            {
                case Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeCoarse:
                    for (int i = 0; i < length; i++) { lines[i + 3] = (i*10).ToString() + "," + voc[i].ToString() + "," + res[i].ToString(); }
                    break;
                case Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeFine:
                    for (int i = 0; i < length; i++) { lines[i + 3] = i.ToString() + "," + voc[i].ToString() + "," + res[i].ToString(); }
                    break;
            }

            List<string> exportFile = new List<string>();
            for (int i=0; i<length+3; i++) { exportFile.Add(lines[i]); }

            StoreDataToCSVFile(exportFile);
        }

        // Changes Made - Change DOC - CheckIndexContent does not exist anymore
        // Creates and saves a new Battery Model from a .csv file to the device
        public void CreateBatteryModelFromCSVFile(System.IO.StreamReader file)
        {
            // extract Voc and res arrays from .csv file
            List<string> fileLines = new List<string>();

            while (!file.EndOfStream)
            {
                fileLines.Add(file.ReadLine());
            }

            file.Close();

            string[] lines = fileLines.ToArray();

            double cap = double.Parse(lines[1].Substring(lines[1].IndexOf("=") + 1, 6).Replace('.', ','));
            double[] voc = new double[lines.Length-3];
            double[] res = new double[lines.Length-3];

            // Create res and Voc arrays
            for (int i = 3; i < lines.Length; i++)
            {
                List<string> separatedValues = new List<string>(lines[i].Split(','));
                voc[i - 3] = double.Parse(separatedValues[1].Replace('.', ','));
                res[i - 3] = double.Parse(separatedValues[2].Replace('.', ','));
            }

            // Sort res and Voc arrays if given with wrong order or with non-monotone values
            Array.Sort(voc);
            Array.Sort(res);
            Array.Reverse(res);

            // Find available index
            index = 1;

            if (CheckIndexContent(index)) { }
            else
            {
                index = FindEmptyIndex();
                Console.WriteLine(index.ToString());
            }

            // Code for n batteries connected (Series or Parallel)
            // terma proxeiros kwdikas
            Console.WriteLine("This is the " + lines[0] + " Battery Model");
            Console.WriteLine("Do you want more same batteries connected in Series or Parallel?");
            if (Console.ReadLine() == "Y")
            {
                Console.WriteLine("How many Batteries?");
                int x = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine("Series or Parallel?");
                if(Console.ReadLine() == "Series") { BatteriesSeries(x, ref voc, ref res); }
                else { BatteriesParallel(x, ref res, ref cap); }
            }

            // Seperate res and resOff array
            double [,] resOffset = SeparateResAndResOffsetArrays(ref res);

            // Check if Battery Model is already saved in the instrument
            if (CheckIfBatteryModelExists(voc, res, cap, resOffset)) return;

            // Save Battery Model to index
            if (index == 0)
            {
                Console.WriteLine("All indexes (1-9) contain other battery models");
                Console.WriteLine("Select index number to be overwritten");
                Console.WriteLine("To abort, press 0");
                index = Convert.ToInt32(Console.ReadLine());
                CheckRange(ref index, 0, 9);
                if (index != 0)
                {
                    driver.Battery.Model.CreateModel("1", Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeFine, index, ref voc, ref res, cap);
                    driver.Battery.Model.Save("1", index);
                }
                else
                {
                    Console.WriteLine("Battery Model not saved.");
                    Console.WriteLine("Battery Resistance Offset Array not saved.");
                    return;
                }
            }
            else
            {
                Console.WriteLine("Battey Model saved on available index " + index.ToString());
                driver.Battery.Model.CreateModel("1", Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeFine, index, ref voc, ref res, cap);
                driver.Battery.Model.Save("1", index);
            }

            driver.Battery.Model.RecallByModelIndex("1", index);
        
            if (!resOffset.Cast<double>().All(x => x == default(double)))
            {
                SaveResistanceOffsetToRepository(resOffset);
            }
        }

        #endregion Battery Model Handling

        #region Simulation

        // ✔
        // Start Simulation Function
        public void StartSimulation()
        {
            EnableOutput();
            List<string> exportFile = SaveData();
            StoreDataToCSVFile(exportFile);
        }

        public void StartSimulation(Keithley2280BatterySimulationModeEnum simMode)
        {
            setMethod(simMode);
        }

        public void StartSimulation(double soc)
        {
            setSOC(soc);
        }

        public void StartSimulation(double soc, double sampInt)
        {
            setSOC(soc);
            SetSampleInterval(sampInt);
        }

        public void StartSimulation(Keithley2280BatterySimulationModeEnum simMode, double soc)
        {
            setMethod(simMode);
            setSOC(soc);
        }

        public void StartSimulation(Keithley2280BatterySimulationModeEnum simMode, double soc, double sampInt)
        {
            setMethod(simMode);
            setSOC(soc);
            SetSampleInterval(sampInt);
            driver.System.DirectIO.WriteString(":BATT:TRAC:CLE");
            driver.Battery.Simulator.ResistanceOffset["1"] = 0;
            EnableOutput();
            List<string> exportFile = SaveData();
            StoreDataToCSVFile(exportFile);
            BatterySimulatorModeClose();
        }

        #endregion Simulation

        #region Save Data and Export Measurements

        public void SaveResistanceOffsetToRepository(double[,] resOffset)
        {
            List<string> exportFile = new List<string>();
            for (int i = 0; i < (resOffset.Length/2); i++)
            {
                exportFile.Add(resOffset[i, 0].ToString().Replace(',','.') + "," + resOffset[i, 1].ToString());
            }

            StoreDataToCSVFile(exportFile);
        }

        // ✔
        // Fetch measurements from the buffer and store them in a List<string>
        public List<string> SaveData()
        {
            List<string> exportFile = new List<string>();
            exportFile.Add("VOC, CURR, SOC, RES, CAP, TST");

            string x = "";
            int i = 1;
            Console.WriteLine("Press any key to stop the simulaton");

            do
            {
                Thread.Sleep(Convert.ToInt32(GetSampleInterval() * 1500));
                driver.System.DirectIO.WriteString(":BATT:TRAC:DATA:SEL? " + i.ToString() + ", " + i.ToString() + ", \"VOLT,CURR,SOC,RES,TST,UNIT\"");
                x = driver.System.DirectIO.ReadString();
                exportFile.Add(x);
                if (i == driver.Measurement.Buffer.BufferSize["1"]) { i = 0; }
                i++;
            } while (!Console.KeyAvailable);

            return exportFile;
        }

        // Create a .csv file with measurements and save it somewhere
        public void StoreDataToCSVFile(List<string> exportFile)
        {
            StackTrace stackTrace = new StackTrace();
            string caller = stackTrace.GetFrame(1).GetMethod().Name;
            Console.WriteLine(caller);

            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string name = "";
            string extension = ".csv";

            switch (caller)
            {
                case "StartSimulation":
                    path += @"\Keithley 2281\Simulation Data\";
                    name = "Test_" + DateTime.Now.Date.ToShortDateString().Replace('.', '_') + '_' + DateTime.Now.Hour.ToString() + 'h' + DateTime.Now.Minute.ToString() + 'm' + extension;
                    break;
                case "SaveResistanceOffsetToRepository":
                    path += @"\Keithley 2281\Battery Models Repository\";
                    name = "ResistanceOffset" + index.ToString() + extension;
                    break;
                case "SaveBatteryModel":
                    path += @"\Keithley 2281\Battery Models Repository\";
                    name = "Model" + index.ToString() + extension;
                    break;
            }
            
            if (!Directory.Exists(path)) { Directory.CreateDirectory(path); }
            path += name;
            System.IO.File.WriteAllLines(path, exportFile);
        }

        #endregion Save Data and Export Measurements

        #endregion Battery Simulator Operation Code
    }
}
