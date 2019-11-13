using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using Ivi.DCPwr.Interop;
using Ivi.Dmm.Interop;
using Ivi.Driver.Interop;
using KeithleyInstruments.Keithley2280.Interop;

namespace diffClasses
{
    public class Keithley2281
    {
        // Initialize new driver component
        public IKeithley2280 driver = new Keithley2280();
        string name = "";
        string operation = "";

        // Constructor with arguments
        public Keithley2281(string logicalName, string operationMode)
        {
            try
            {
                driver.Initialize(logicalName, true, false, "");
                name = logicalName;
                operation = operationMode;
                Console.WriteLine(driver.Identity.InstrumentManufacturer + " , " + driver.Identity.InstrumentModel);
                chooseOperationMode(operationMode);
            }
            catch (Exception e)
            {
                Console.WriteLine("Not a Valid Instrument Resource Name");
            }
        }

        // Choose Operation Mode
        public void chooseOperationMode(string operationMode)
        {
            Console.WriteLine("skata");
            if (driver.Initialized)
            {
                //channel = chooseChannel();

                driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeEntry; // correct is [channel.ToString()] error -114 when =2

                switch (operationMode)
                {
                    case "Power Supply":
                        PowerSupply ps = new PowerSupply(driver);
                        ps.powerSupplyMode();
                        break;
                    case "Battery Test":
                        Test battery = new Test(driver);
                        battery.batteryTestMode();
                        break;
                    case "Battery Simulator":
                        Simulator bat = new Simulator(name, operation, driver);
                        bat.batterySimulatorMode();
                        break;
                    default:
                        Console.WriteLine("Wrong String Value");
                        Console.WriteLine("Decide between Power Supply , Battery Test and Battery Simulator");
                        operationMode = Console.ReadLine();
                        chooseOperationMode(operationMode);
                        break;
                }
            }
        }

        // Check Range Function
        public int checkRange(int val, int margin1, int margin2)
        {
            while (val < margin1 || val > margin2)
            {
                Console.WriteLine("Error. Please try again");
                val = Convert.ToInt32(Console.ReadLine());
            }
            return val;
        }

        // Choose Channel
        public int chooseChannel()  // must include ref value
        {
            int channel = 0;

            Console.WriteLine("Please choose desired channel:");
            Console.WriteLine("1. Channel 1");
            Console.WriteLine("2. Channel 2");

            channel = checkRange(Convert.ToInt32(Console.ReadLine()), 1, 2);

            return channel;
        }

        // Close Connection with Device
        public void closeConnection()
        {
            if (driver.Initialized)
            {
                driver.Close();
            }
            Console.WriteLine("Session terminated succesfully");
        }

        // Return device errors - NOT YET USED!
        public void deviceErrors()
        {
            int code = 0;
            string mess = "";
            driver.Utility.ErrorQuery(ref code, ref mess);
            Console.WriteLine("Error code: " + code.ToString());
            Console.WriteLine("Error Message: " + mess);
        }

        public class PowerSupply : Keithley2281
        {
            #region Power Supply Operation Code

            IKeithley2280 driver = new Keithley2280();

            public PowerSupply(IKeithley2280 x) : base("a", "b")
            {
                driver = x;
            }

            // All Power Supply Mode Functions and Methods
            public void powerSupplyMode()
            {
                driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModePowerSupply;
                Thread.Sleep(3000);
                powerSupplyModeClose();
            }

            public void powerSupplyModeClose()
            {
                Console.WriteLine("Choose new operation mode");
                string operationMode = Console.ReadLine();
                chooseOperationMode(operationMode);
            }

            #endregion Power Supply Operation Code
        }

        public class Test : Keithley2281 // : Battery
        {
            #region Battery Test Operation Code

            IKeithley2280 driver = new Keithley2280();

            public Test(IKeithley2280 x) : base("", "")
            {
                driver = x;
            }

            // All Battery Test Mode Functions and Methods
            public void batteryTestMode()
            {
                driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatteryTest;
                Thread.Sleep(3000);
                batteryTestModeClose();
            }

            public void batteryTestModeClose()
            {
                Console.WriteLine("Choose new operation mode");
                string operationMode = Console.ReadLine();
                chooseOperationMode(operationMode);
            }

            #endregion Battery Test Operation Code
        }

        public class Simulator : Keithley2281 // : Battery
        {
            #region Battery Simulator Operation Code

            IKeithley2280 driver = new Keithley2280();

            public Simulator(string name, string operation, IKeithley2280 x) : base(name, operation)
            {
                driver = x;
            }


            // Initiate Battery Simulator mode
            public void batterySimulatorMode()
            {
                Console.WriteLine("MPHKA");
                driver.System.OperationMode["1"] = Keithley2280SystemOperationModeEnum.Keithley2280SystemOperationModeBatterySimulator;
                Console.WriteLine("MPHKA");
                driver.Battery.Simulator.SimulationMode["1"] = Keithley2280BatterySimulationModeEnum.Keithley2280BatterySimulationModeDynamic;
                Console.WriteLine("MPHKA");
                saveBatteryModelToUSB(10);
                Thread.Sleep(3000);
                Console.WriteLine("MPJKA");
                batterySimulatorModeClose();
            }

            #region TO BE FILLED

            // Check if desired battery model already exists on device
            // DE MAS NOIZEI KA TOSO AMA DEN YPARXEI!!!!
            public void checkIfBatteryModelExists(string batteryModelName)
            {
                for (int i = 1; i < 15; i++)
                {

                }
            }

            // Choose a battery model
            public void chooseBatteryModel(string batteryType)
            {

            }

            // via index
            public void loadBatteryModel(int index)
            {
                if (index < 15 && index > 0)
                {
                    driver.Battery.Model.RecallByModelIndex("1", index);
                }
                else
                {
                    Console.WriteLine("Index out of Bounds");
                }
            }

            public void checkFileContent(System.IO.StreamReader file)
            {

            }

            public void loadBatteryModel(System.IO.StreamReader file)
            {

            }

            public void saveBatteryModelToDevice(int index)
            {
                if (index < 15 && index > 0)
                {
                    driver.Battery.Model.Save("1", index);
                }
                else
                {
                    Console.WriteLine("Index out of Bounds");
                }
            }

            public void saveBatteryModelToUSB(int index) //prepei na checkarw an uparxei idio onoma! // to USB prepei na einai panw sto mhxanhma, oxi sto PC
            {
                Random rand = new Random();
                rand.Next(16000);
                string fileName = "batteryModel" + rand.ToString();
                driver.Battery.Model.SaveToUSB("1", index, fileName);
            }

            // NEEDS WORK!
            public void saveBatteryModel()
            {
                //Save csv file
                double[] voc = new double[10];
                double[] res = new double[10];
                double cap = 0.1;

                // COMMAND CAUSES ERROR
                driver.Battery.Model.QueryModel("1", Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeFine, 10, ref voc, ref res, cap);

                string[] lines = new string[11];
                lines[0] = "Voc,Res";

                for (int i = 0; i < 10; i++)
                {
                    Console.WriteLine(voc[i]);
                    Console.WriteLine(res[i]);
                    lines[i + 1] = voc[i].ToString() + "," + res[i].ToString();
                }

                Console.Write(lines);

                // SaveDialog Method

                Stream str = new MemoryStream();
                var writer = new StreamWriter(str);
                writer.Write(lines);
                SaveFileDialog save = new SaveFileDialog();

                save.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
                save.FilterIndex = 2;
                save.RestoreDirectory = true;

                if (save.ShowDialog() == DialogResult.OK)
                {
                    if ((str = save.OpenFile()) != null)
                    {
                        str.Close();
                    }
                }

                System.IO.File.WriteAllLines(@"C:\Users\prakHW\Desktop\test.csv", lines);
            }

            public void createBatteryModelFromMDLFile(System.IO.StreamReader file)
            {

            }

            public void createBatteryModelFromCSVFile(System.IO.StreamReader file)
            {
                List<string> fileLines = new List<string>();

                while (!file.EndOfStream)
                {
                    fileLines.Add(file.ReadLine());
                }

                string[] lines = fileLines.ToArray();
                double[] voc = new double[lines.Length];
                double[] res = new double[lines.Length];

                for (int i = 0; i < lines.Length; i++)
                {
                    List<string> separatedValues = new List<string>(lines[i].Split(','));
                    voc[i] = double.Parse(separatedValues[0]);
                    res[i] = double.Parse(separatedValues[1]);
                }

                driver.Battery.Model.CreateModel("1", Keithley2280BatteryModelTypeEnum.Keithley2280BatteryModelTypeFine, 1, ref voc, ref res, 0.1); // random values
            }

            // Configure Battery Situation
            public void configureBatterySim(Keithley2280BatterySimulationModeEnum a)
            {
                driver.Battery.Simulator.SimulationMode["1"] = Keithley2280BatterySimulationModeEnum.Keithley2280BatterySimulationModeDynamic;
            }

            // Overload(s) of Battery Situation Config
            public void configureBatterySim()
            {

            }

            #endregion TO BE FILLED

            #region SETTERS AND GETTERS (STATUS NEED TO BE FILLED)

            // Method getter and setter
            public void setMethod(Keithley2280BatterySimulationModeEnum method)
            {
                driver.Battery.Simulator.SimulationMode["1"] = method;
            }

            public Keithley2280BatterySimulationModeEnum getMethod()
            {
                return driver.Battery.Simulator.SimulationMode["1"];
            }

            // VOC getter and setter
            public void setVOC(double voc)
            {
                driver.Battery.Simulator.OpenCircuitVoltage["1"] = voc;
            }

            public double getVOC()
            {
                return driver.Battery.Simulator.OpenCircuitVoltage["1"];
            }

            // SOC getter and setter
            public void setSOC(double soc)
            {
                driver.Battery.Simulator.StateOfCharge["1"] = soc;
            }

            public double getSOC()
            {
                return driver.Battery.Simulator.StateOfCharge["1"];
            }

            // Full Voltage getter and setter
            public void setFullV(double fullV)
            {
                driver.Battery.Simulator.OpenCircuitVoltageFull["1"] = fullV;
            }

            public double getFullV()
            {
                return driver.Battery.Simulator.OpenCircuitVoltageFull["1"];
            }

            // Empty Voltage getter and setter
            public void setEmptyV(double emptyV)
            {
                driver.Battery.Simulator.OpenCircuitVoltageEmpty["1"] = emptyV;
            }

            public double getEmptyV()
            {
                return driver.Battery.Simulator.OpenCircuitVoltageEmpty["1"];
            }

            // I Limit getter and setter
            public void setILimit(double iLimit)
            {
                driver.Battery.Simulator.CurrentLimit["1"] = iLimit;
            }

            public double getILimit()
            {
                return driver.Battery.Simulator.CurrentLimit["1"];
            }

            // Capacity Limit getter and setter
            public void setCapacityLimit(double capLimit)
            {
                driver.Battery.Simulator.CapacityLimit["1"] = capLimit;
            }

            public double getCapacityLimit()
            {
                return driver.Battery.Simulator.CapacityLimit["1"];
            }

            // OVP getter and setter
            public void setOVP(double ovp)
            {
                if (ovp < 0.5 || ovp > 21)
                {
                    Console.WriteLine("OVP Limit: 0.5 - 21 V \nValue out Of Range \nValue set to maxValue = 21 V");
                    setOVP(Convert.ToDouble(Console.ReadLine()));
                }
                else
                {
                    driver.Battery.Simulator.OVPLimit["1"] = ovp;
                }
            }

            public double getOVP()
            {
                return driver.Battery.Simulator.OVPLimit["1"];
            }

            // OCP getter and setter
            public void setOCP(double ocp)
            {
                if (ocp < 0.1 || ocp > 6.1)
                {
                    Console.WriteLine("OCP Limit: 0.1 - 6.1 A \nValue out of Range \nValue set to maxValue = 6.1 A");
                    setOCP(Convert.ToDouble(Console.ReadLine()));
                }
                else
                {
                    driver.Battery.Simulator.OCPLimit["1"] = ocp;
                }
            }

            public double getOCP()
            {
                return driver.Battery.Simulator.OCPLimit["1"];
            }

            // Resistance Offset getter and setter
            public void setResistanceOffset(double resOffset)
            {
                driver.Battery.Simulator.ResistanceOffset["1"] = resOffset;
            }

            public double getResistanceOffset()
            {
                return driver.Battery.Simulator.ResistanceOffset["1"];
            }

            /*
            // Status getter and setter
            public void setStatus(Keithley2280BatteryOutputStateEnum state)
            {

                driver.Battery.Simulator. = Keithley2280BatteryOutputStateEnum.;
            }

            public Keithley2280BatteryOutputStateEnum getStatus()
            {
                return driver.Battery.OutputState["1"];
            }
            */

            // Get Resistance
            public double getResistance()
            {
                return driver.Battery.Simulator.Resistance["1"];
            }

            // Get Current
            public double getCurrent()
            {
                return driver.Battery.Simulator.Current["1"];
            }

            // Get Capacity
            public double getCapacity()
            {
                return driver.Battery.Simulator.Capacity["1"];
            }

            // Get Terminal Voltage
            public double getTerminalVoltage()
            {
                return driver.Battery.Simulator.TerminalVoltage["1"];
            }

            #endregion SETTERS AND GETTERS (STATUS NEED TO BE FILLED)

            // Close Battery Simulator mode
            public void batterySimulatorModeClose()
            {
                Console.WriteLine("Write next desired operation mode");
                string operationMode = Console.ReadLine();
                this.chooseOperationMode(operationMode);
            }

            #endregion Battery Simulator Operation Code
        }
    }
}
