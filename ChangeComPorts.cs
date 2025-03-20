// Author: Sawyer Zock
// This program reassigns the com port of the default windows Communications Port from COM1 to COM9, then assigns the USB Serial Port for a Topaz Signature Pad to COM1.
// After this program is run the Topaz device must be disconnected then reconnected, either physically or through CMD/Powershell like below.
//
//pnputil /restart-device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
//or
//pnputil /disable-device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
//pnputil / enable - device "FTDIBUS\VID_0403+PID_6001+TOPAZBSBA\0000"
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


Console.WriteLine("Com Port script initilized...");
Console.WriteLine("Com Ports before change:");
String[] allPorts = System.IO.Ports.SerialPort.GetPortNames();
foreach (string port in allPorts)
{
    Console.WriteLine(port);
}

Console.WriteLine("Releasing COM ports.");




//Release active com parts list in current control set, byte 0 is com1 - 8, byte 1 is 9-16. We want both to be 0 to release the com ports.
RegistryKey currentControlArbiterClear = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter", true);
if (currentControlArbiterClear != null)
{
    byte[] data = (byte[])currentControlArbiterClear.GetValue("ComDB");
    if (data != null)
    {
        data[0] = (byte)0x00;
        data[1] = (byte)0x00;
        currentControlArbiterClear.SetValue("ComDB", data);
    }
    currentControlArbiterClear.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter is Null, something is VERY wrong.");
}

//Releasing COM ports for controlset1
RegistryKey controlOneArbiterClear = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Control\\COM Name Arbiter", true);
if (controlOneArbiterClear != null)
{
    byte[] data = (byte[])controlOneArbiterClear.GetValue("ComDB");
    if (data != null)
    {
        data[0] = (byte)0x00;
        data[1] = (byte)0x00;
        controlOneArbiterClear.SetValue("ComDB", data);
    }
    controlOneArbiterClear.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Control\\COM Name Arbiter is Null, something is VERY wrong.");
}

Console.WriteLine("COM ports now released.");


//Sets port to COM9 in current control set
RegistryKey currentControlDefaultPort = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2\\Device Parameters", true);
if (currentControlDefaultPort != null)
{
    currentControlDefaultPort.SetValue("PortName", "COM9", RegistryValueKind.String);
    currentControlDefaultPort.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2\\Device Parameters is Null.");
}

//Sets port to COM9 in control set 1
RegistryKey controlOneDefaultPort = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2\\Device Parameters", true);
if (controlOneDefaultPort != null)
{
    controlOneDefaultPort.SetValue("PortName", "COM9", RegistryValueKind.String);
    controlOneDefaultPort.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2\\Device Parameters is Null.");
}


//Sets friendly name to COM9 in current control set
RegistryKey currentControlSIOBUAR2Key = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2", true);
if (currentControlSIOBUAR2Key != null)
{
    currentControlSIOBUAR2Key.SetValue("FriendlyName", "Communications Port (COM9)", RegistryValueKind.String);
    currentControlSIOBUAR2Key.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2 is Null.");
}

//Sets friendly name to COM9 in control set 1
RegistryKey controlOneSIOBUAR2Key = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2", true);
if (controlOneSIOBUAR2Key != null)
{
    controlOneSIOBUAR2Key.SetValue("FriendlyName", "Communications Port (COM9)", RegistryValueKind.String);
    controlOneSIOBUAR2Key.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2 is Null.");
}

//Extracts the data contained within the default windows COM1 key, then renames the key while preserving the original data contained within it.

RegistryKey currentControlArbiterKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2", true);
if (currentControlArbiterKey != null)
{

    String regDataTwo = currentControlArbiterKey.GetValue("COM9").ToString();
    if (regDataTwo != null)
    {
        currentControlArbiterKey.SetValue("COM9", regDataTwo, RegistryValueKind.String);
    }

    currentControlArbiterKey.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Enum\\ACPI\\PNP0501\\SIOBUAR2 is Null.");
}


//Extracts the data contained within the default windows COM1 key, then renames the key while preserving the original data contained within it.
RegistryKey controlOneArbiterKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2", true);
if (controlOneArbiterKey != null)
{

    String regData = controlOneArbiterKey.GetValue("COM9").ToString();
    if (regData != null)
    {
        controlOneArbiterKey.SetValue("COM9", regData, RegistryValueKind.String);
    }
    controlOneArbiterKey.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\ACPI\\PNP0501\\SIOBUAR2 is Null.");
}

Console.WriteLine("Default windows Communication Port is now assigned to COM9.");
//End COM9 changes


Console.WriteLine("Changing Topaz Communication Port to COM1");
//Begin COM1 changes
//Sets Topaz port to COM1 in current control set
RegistryKey currentControlTopazPort = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", true);
if (currentControlTopazPort != null)
{
    currentControlTopazPort.SetValue("PortName", "COM1", RegistryValueKind.String);
    currentControlTopazPort.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters is Null, verify Topaz is installed.");
}

//Sets Topaz port to COM1 in control set 1
RegistryKey controlOneTopazPort = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", true);
if (controlOneTopazPort != null)
{
    controlOneTopazPort.SetValue("PortName", "COM1", RegistryValueKind.String);
    controlOneTopazPort.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters is Null, verify Topaz is installed.");
}

//Sets friendly name to COM1 in current control set
RegistryKey currentControlFTDIBUSKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000", true);
if (currentControlFTDIBUSKey != null)
{
    currentControlFTDIBUSKey.SetValue("FriendlyName", "USB Serial Port (COM1)", RegistryValueKind.String);
    currentControlFTDIBUSKey.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000 is Null, verify Topaz is installed.");
}

//Sets friendly name to COM1 in control set 1
RegistryKey controlOneFTDIBUSKey = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000", true);
if (controlOneFTDIBUSKey != null)
{
    controlOneFTDIBUSKey.SetValue("FriendlyName", "USB Serial Port (COM1)", RegistryValueKind.String);
    controlOneFTDIBUSKey.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000 is Null, verify Topaz is installed.");
}




Console.WriteLine("Setting Topaz advanced options...");
//This enables four options in Topaz's advanced settings: Cancel if power off, Event on surprise removal, Send RTS on close, and Disable modem ctrl at startup.
//Current control set
RegistryKey currentControlAdvanced = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", true);
if (currentControlAdvanced != null)
{
    byte[] data = (byte[])currentControlAdvanced.GetValue("ConfigData");
    if (data != null)
    {
        data[0] = (byte)0x79;
        data[1] = (byte)0x02;
        currentControlAdvanced.SetValue("ConfigData", data);
    }
    currentControlAdvanced.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters is Null, verify Topaz is installed. ");
}


//This enables four options in Topaz's advanced settings: Cancel if power off, Event on surprise removal, Send RTS on close, and Disable modem ctrl at startup.
//Control set 1
RegistryKey controlOneAdvanced = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters", true);
if (controlOneAdvanced != null)
{
    byte[] data = (byte[])controlOneAdvanced.GetValue("ConfigData");
    if (data != null)
    {
        data[0] = (byte)0x79;
        data[1] = (byte)0x02;
        controlOneAdvanced.SetValue("ConfigData", data);
    }
    controlOneAdvanced.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Enum\\FTDIBUS\\VID_0403+PID_6001+TOPAZBSBA\\0000\\Device Parameters is Null, verify Topaz is installed. ");
}

Console.WriteLine("Topaz advanced settings changed.");


//Change active com parts list in current control set, byte 0 is com1 - 8, byte 1 is 9-16. Ports 1 and 9 will be active, so we set the first two bytes to 01 01
RegistryKey currentControlArbiter = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter", true);
if (currentControlArbiter != null)
{
    byte[] data = (byte[])currentControlArbiter.GetValue("ComDB");
    if (data != null)
    {
        data[0] = (byte)0x01;
        data[1] = (byte)0x01;
        currentControlArbiter.SetValue("ComDB", data);
    }
    currentControlArbiter.Close();
}
else
{
    Console.WriteLine("SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter is Null, something is VERY wrong.");
}

//Change active com parts list in control set 1, byte 0 is com1 - 8, byte 1 is 9-16. Ports 1 and 9 will be active, so we set the first two bytes to 01 01
RegistryKey controlOneArbiter = Registry.LocalMachine.OpenSubKey("SYSTEM\\ControlSet001\\Control\\COM Name Arbiter", true);
if (controlOneArbiter != null)
{
    byte[] data = (byte[])controlOneArbiter.GetValue("ComDB");
    if (data != null)
    {
        data[0] = (byte)0x01;
        data[1] = (byte)0x01;
        controlOneArbiter.SetValue("ComDB", data);
    }
    controlOneArbiter.Close();
}
else
{
    Console.WriteLine("SYSTEM\\ControlSet001\\Control\\COM Name Arbiter is Null, something is VERY wrong.");
}

Console.WriteLine("COM ports now reserved.");


//Sets serial com devicemap to correct values
RegistryKey deviceMapSerialComKey = Registry.LocalMachine.OpenSubKey("HARDWARE\\DEVICEMAP\\SERIALCOMM", true);
if (deviceMapSerialComKey != null)
{
    deviceMapSerialComKey.SetValue("\\Device\\Serial0", "COM9", RegistryValueKind.String);
    deviceMapSerialComKey.SetValue("\\Device\\VCP0", "COM1", RegistryValueKind.String);
    deviceMapSerialComKey.Close();
}
else
{
    Console.WriteLine("HARDWARE\\DEVICEMAP\\SERIALCOMM Registry Key is Null, Windows must have changed the registry location for Device Map");
    Console.WriteLine("Contact Sawyer Zock to update this script.");
}

Console.WriteLine("COM ports now reserved.");

//Printing all COM ports to show change.

Console.WriteLine("Ports after change:");
allPorts = System.IO.Ports.SerialPort.GetPortNames();
foreach (string port in allPorts)
{
    Console.WriteLine(port);
}

Console.WriteLine("\nAll registry changes comeplete, now resetting Topaz device.");
//Resetting Topaz device
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
    Console.WriteLine("Failed to run pnputil.exe, reseat Topaz manually.");
}

Console.ReadLine();
Environment.Exit(0);


