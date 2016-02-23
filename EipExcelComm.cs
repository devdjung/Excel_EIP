using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using AdvancedHMIDrivers;

namespace EipExcelComm_lib
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class EipExcelComm
    {
        private AdvancedHMIDrivers.EthernetIPforCLXComm Plc;

        public EipExcelComm()
        {
            Plc = new EthernetIPforCLXComm();

        }
        public void setPlcIPAddress(string ip)
        {
            Plc.IPAddress = ip;
        }
        public void setPlcSlot(int slot)
        {
            Plc.ProcessorSlot = slot;
        }
        public string getData(string tag)
        {
            string data = "empty";
            try
            {
                data = Plc.ReadAny(tag);
            }
            catch(Exception exc)
            {
                data = "E: " + exc;
            }
            return data;
        }
        public void close()
        {
            Plc.Dispose();
        }
    }
}
