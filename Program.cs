// Looks like there is a bug with the Printer that holds the previous POS command responce in the buffer and tranmit it when a new status command request is sent. Therefore it is better to separete the Printer Statuses and Printer Cutters Number checkers in two individuals exe programs. 
// Oracle Simphony keeps holding the port connection some times, even when the Label Printer is not in use. Therefore, we need to restart the Simphony application to be able to get a responce from the Printer.
// It would be advisible to send the same command a sencond time case there is no response from printer first time.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.ComponentModel;
using System.Net;
using System.Runtime.InteropServices;
using System.Net.Mail;
using Microsoft.VisualBasic;

namespace PrinterMonitoringSample
{
    class ESC_POS_Command
    {
        [DllImport("iphlpapi.dll", ExactSpelling = true)]
        static extern int SendARP(int destIp, int srcIP, byte[] macAddr, ref uint physicalAddrLen);
        static void Main(string[] args)
        {
            bool isHeader = true;
            bool Reachable;
            string mac_Address = "";
            string Response_status = "";
            foreach (var line in File.ReadLines("Devices_IPs.txt"))
            {
                if (isHeader)
                {
                    isHeader = false;
                    continue;
                }
                var fields = line.Split(","); //Console.WriteLine($"IP Address: {fields[0]} Time out: {fields[1]} Command: {fields[2]}");
                string ipAddress = fields[0];
                int timeOut = Int32.Parse(fields[1]);
                string deviceName = fields[2];
                //Console.WriteLine($"IP Address: {fields[0]} Time out: {fields[1]} Command: {fields[2]}");

                try
                {
                    if (new FileInfo("Printer_status.log").Length > 10000) File.Delete("Printer_status.log"); // Keep the Log file size below 10K bytes
                }
                catch (FileNotFoundException e)
                {
                    //Console.WriteLine(e.ToString());
                }

                try
                {
                    if (new FileInfo("Ping_status.log").Length > 10000) File.Delete("Ping_status.log"); // Delete the Log file is over 10K bytes
                }
                catch (FileNotFoundException e)
                {
                    //Console.WriteLine(e.ToString());
                }

                Reachable = PingHost(ipAddress, timeOut, deviceName);
                //if (Reachable == true) PrinterStatus(ipAddress, deviceName);
                if (Reachable == true && deviceName.ToUpperInvariant().Contains("LABEL")) PrinterCuts(ipAddress, deviceName);
            }
            void PrinterCuts(string ipAddress, string DeviceName)
            {
                //string Response_status = "Test1";
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = String.Format("/c plink.exe {0} -P {1} -raw -v < cmd_cutter.txt", ipAddress, "9100"), // Check how many cutters
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true,
                    RedirectStandardInput = false,
                    CreateNoWindow = true,
                };
                Process proc = Process.Start(psi);
                Response_status = proc.StandardOutput.ReadToEnd(); // rediredct the Plink Output to a String Variable
                if (proc != null && proc.HasExited != true)
                {
                    proc.WaitForExit();
                }
                if (Response_status == "")
                {
                    proc = Process.Start(psi); //Try to get the printer responce one more time.
                    Response_status = proc.StandardOutput.ReadToEnd();
                    if (proc != null && proc.HasExited != true)
                    {
                        proc.WaitForExit();
                    }
                    if (Response_status == "")
                    {
                        WriteLog_PrinterResponse("No cutter response from Printer", ipAddress, DeviceName); // If no reply to the ESC/POS command from the printer.
                    }
                    else
                    {
                        String ResponseStatus = Response_status.Replace("_", "");
                        ResponseStatus = ResponseStatus.Replace("\0", "");  //Used to remove the "NULL" character at the end of the Number of cuts response, only visible with notepad++, that prevents the CSM to capture the events
                        WriteLog_PrinterResponse("Number of cuts = " + ResponseStatus, ipAddress, DeviceName);
                    }
                }
                else
                {
                    String ResponseStatus = Response_status.Replace("_", "");
                    ResponseStatus = ResponseStatus.Replace("\0", "");  //Used to remove the "NULL" character at the end of the Number of cuts response, only visible with notepad++, that prevents the CSM to capture the events
                    WriteLog_PrinterResponse("Number of cuts = " + ResponseStatus, ipAddress, DeviceName);
                }
            }
            bool PingHost(string IP_Address, int TimeOut, string DeviceName)
            {
                bool Reachable = false;
                Ping? pinger = null;
                try
                {
                    pinger = new Ping();
                    PingReply reply = pinger.Send(IP_Address, TimeOut);
                    Reachable = reply.Status == IPStatus.Success;
                    if (Reachable == true)
                    {
                        mac_Address = getMAC_Address(IP_Address);
                    }
                }
                catch (PingException)
                {
                    // Discard PingExceptions and return false;
                }
                finally
                {
                    if (pinger != null)
                    {
                        pinger.Dispose();
                    }
                }
                WriteLog_PingResponse(IP_Address, Reachable, DeviceName, mac_Address);
                return Reachable;
            }

            static String getMAC_Address(string ipAddr)
            {
                IPAddress dst = IPAddress.Parse(ipAddr);

                byte[] macAddr = new byte[6];
                uint macAddrLen = (uint)macAddr.Length;

                if (SendARP(BitConverter.ToInt32(dst.GetAddressBytes(), 0), 0, macAddr, ref macAddrLen) != 0)
                    throw new InvalidOperationException("SendARP failed.");

                string[] str = new string[(int)macAddrLen];
                for (int i = 0; i < macAddrLen; i++)
                    str[i] = macAddr[i].ToString("x2");

                String MAC_Address = string.Join("-", str);
                return MAC_Address;
            }

            static void WriteLog_PrinterResponse(string text, string ipAddr, string DeviceName)
            {
                string log = System.DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss:fff]");
                log += DeviceName + " : ";
                log += text;
                log += " (IP : " + ipAddr + ")";
                log += "\r\n";
                using FileStream stream = File.Open("Printer_status.log", FileMode.Append, (FileAccess)FileShare.Write); //Trying to avoid the error 1026 using FileStream;
                {
                    var bytes = Encoding.UTF8.GetBytes(log); // Convert to Bytes to be able to use FileStream
                    stream.Write(bytes, 0, bytes.Length);
                }
            }
            static void WriteLog_PingResponse(string ipAddr, bool Reachable, string DeviceName, string mac_Address)
            {
                string log = System.DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss:fff] ");
                log += DeviceName + ", ";
                log += "MAC Address: ";
                log += mac_Address + ", IP address: ";
                log += ipAddr + " reachable = " + Reachable;

                log += "\r\n";
                using FileStream stream = File.Open("Ping_status.log", FileMode.Append, (FileAccess)FileShare.Write); //Trying to avoid the error 1026 using FileStream;
                {
                    var bytes = Encoding.UTF8.GetBytes(log); // Convert to Bytes to be able to use FileStream
                    stream.Write(bytes, 0, bytes.Length);
                }
            }
        }

    }
}
