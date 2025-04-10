// Author: Sawyer Zock
// This program reassigns the com port of the default windows Communications Port from COM1 to COM9, then assigns the USB Serial Port for a Topaz Signature Pad to COM1.
// After this program is run the Topaz device must be disconnected then reconnected, either physically or through CMD/Powershell like below.
//
//pnputil /restart-device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
//or
//pnputil /disable-device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
//pnputil /enable - device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
//
//I have integrated this at the end of the program.
// Note: Reassigning COM ports programatically should generally not be done, COM ports should be set by the drivers.
// This is for cases when you do not have the access or authority to properly set up devices to use any port dynamically rather than being statically assigned. 

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO.Ports;
using System.Text.Json.Serialization;
using Microsoft.Win32;


Console.WriteLine("Beginning Com Port script...");


//The port the default windows serial port will be set to.
string targetPortDefault = "COM9";
//The port the Topaz device will be set to.
string targetPortTopaz = "COM1";


Console.WriteLine("Com Ports before change:");
printPorts();

Console.WriteLine("Releasing COM ports.");
setComDB("SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter", (byte)0x00, (byte)0x00);
setComDB("SYSTEM\\ControlSet001\\Control\\COM Name Arbiter", (byte)0x00, (byte)0x00);

Console.WriteLine("Setting windows default serial port.");
setPortName("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2\\Device Parameters", targetPortDefault);
setPortName("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2\\Device Parameters", targetPortDefault);

Console.WriteLine("Setting windows default serial port friendly name.");
setFriendlyName("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2", "Communications Port (" + targetPortDefault + ")");
setFriendlyName("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2", "Communications Port (" + targetPortDefault + ")");

Console.WriteLine("Setting Topaz serial port.");
setPortName("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", "COM1");
setPortName("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", "COM1");

Console.WriteLine("Setting Topaz serial port friendly name.");
setFriendlyName("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000", "Communications Port (" + targetPortTopaz + ")");
setFriendlyName("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000", "Communications Port (" + targetPortTopaz + ")");

Console.WriteLine("Setting Topaz advanced device settings.");
setTopazAdvanced("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", (byte)0x79, (byte)0x02);
setTopazAdvanced("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", (byte)0x79, (byte)0x02);


Console.WriteLine("Reserving COM ports.");
setComDB("SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter", (byte)0x01, (byte)0x01);
setComDB("SYSTEM\\ControlSet001\\Control\\COM Name Arbiter", (byte)0x01, (byte)0x01);

Console.WriteLine("Updating device map.");
setGenericKey("HARDWARE\\DEVICEMAP\\SERIALCOMM", "\\Device\\Serial0", "COM9", RegistryValueKind.String);
setGenericKey("HARDWARE\\DEVICEMAP\\SERIALCOMM", "\\Device\\VCP0", "COM1", RegistryValueKind.String);

Console.WriteLine("Com Ports after change:");
printPorts();

resetTopaz();
exit();


static void printPorts()
{
    String[] allPorts = System.IO.Ports.SerialPort.GetPortNames();
    foreach (string port in allPorts)
    {
        Console.WriteLine(port);
    }
}



static void setPortName(string keyPath, string comPort)
{
    setGenericKey(keyPath, "PortName", comPort, RegistryValueKind.String);
}
static void setFriendlyName(string keyPath, string name)
{
    setGenericKey(keyPath, "FriendlyName", name, RegistryValueKind.String);
}


static void setGenericKey(string keyPath, string keyName, string? keyValue, Microsoft.Win32.RegistryValueKind valueKind)
{
    //Sets port to COM9 in current control set
    RegistryKey? regKey = Registry.LocalMachine.OpenSubKey(keyPath, true);
    if (regKey != null)
    {
        if (keyValue != null)
        {
            regKey.SetValue(keyName, keyValue, valueKind);
        }
        else
        {
            Console.WriteLine("Cannot write a null value to a registry key, check the data passed into the function call. ");
        }
        regKey.Close();
    }
    else
    {
        Console.WriteLine(keyPath + " is Null.");
    }
}


static void setComDB(string keyPath, byte byteOne, byte byteTwo)
{   //Sets which com ports are reserved.
    //Byte 0 is com1 - 8, byte 1 is 9-16 e.g.:
    //Set both to be 0x00 to release the com ports, 0x01 0x01 to reserve ports 1 and 9, 0x03 0x02 to reserve ports 3 and 10.
    setGenericBinaryArray(keyPath, "ComDB", byteOne, byteTwo);
}


static void setTopazAdvanced(string keyPath, byte byteOne, byte byteTwo)
{   //This enables four options in Topaz's advanced settings: Cancel if power off, Event on surprise removal, Send RTS on close, and Disable modem ctrl at startup.
    setGenericBinaryArray(keyPath, "ConfigData", byteOne, byteTwo);
}

static void setGenericBinaryArray(string keyPath, string keyName, byte byteOne, byte byteTwo)
{//This only changes the first two bytes of a binary array, a new function would need to be made to change bytes beyond the first two.

    RegistryKey? regKey = Registry.LocalMachine.OpenSubKey(keyPath, true);
    if (regKey != null)
    {
        byte[]? data = (byte[]?)regKey.GetValue(keyName);
        if (data != null)
        {
            data[0] = byteOne;
            data[1] = byteTwo;
            regKey.SetValue(keyName, data);
        }
        else
        {//This should never be null unless something damaged the registry or if the registry path changes in windows 12+.
            Console.WriteLine(keyName + " data is null at " + keyPath);
        }
        regKey.Close();

    }
    else
    {//This should never be null unless something damaged the registry or if the registry path changes in windows 12+.
        Console.WriteLine(keyPath + " is Null, registry key does not exist.");
    }
}


static void resetTopaz()
{
    Console.WriteLine("\nTrying to reset Topaz device.");
    //To do this in command line, run the command below:
    //pnputil /restart-device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
    try
    {
        var pnpCall = new System.Diagnostics.Process();
        System.Diagnostics.ProcessStartInfo pnpArgs = new System.Diagnostics.ProcessStartInfo();
        pnpArgs.FileName = @"pnputil.exe";
        pnpArgs.Arguments = "/restart-device \"FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\"";
        pnpArgs.Verb = "runas";
        pnpCall.StartInfo = pnpArgs;
        pnpCall.Start();
    }
    catch
    {
        Console.WriteLine("Failed to run pnputil.exe, reset Topaz manually by unplugging and replugging it.");
    }
    Console.ReadLine();
}

static void exit()
{
    Environment.Exit(0);
}

