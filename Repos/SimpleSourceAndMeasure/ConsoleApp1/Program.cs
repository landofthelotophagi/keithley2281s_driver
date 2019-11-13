using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keithley2281_Handler_Version1;
using KeithleyInstruments.Keithley2280.Interop;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Keithley2281 anti = new Keithley2281("USB0::0x05E6::0x2281::4380228::INSTR");
            anti.BatterySimulatorMode(@"C:\Users\prakHW\Desktop\vartaCR2450_4x.csv");
            anti.StartSimulation(Keithley2280BatterySimulationModeEnum.Keithley2280BatterySimulationModeStatic, 7, 0.048);
            Console.ReadKey();
        }
    }
}
