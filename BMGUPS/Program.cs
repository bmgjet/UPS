using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Threading;

namespace BMGUPS
{
	class Program
	{
		public static float PowerFactor = 0.8f;
		public static ulong w = 0;
		public static int Batterys = 6;
		public static int UPSSize = 2400;
		public static float CostKWh = 0.33f;

		static void Main(string[] args)
		{
			string FileDump = string.Empty;
			bool DumpCSV = false;
			if(args.Length > 0)
            {
				DumpCSV = true;
				FileDump = args[0];
			}
			Console.CancelKeyPress += (sender, e) =>
			{
				e.Cancel = true;
				SaveData();
				Environment.Exit(0);
			};
			AppDomain.CurrentDomain.ProcessExit += (sender, e) =>{SaveData();};
			if (File.Exists("KWh.txt")){ulong.TryParse(File.ReadAllText("KWh.txt"), out w);}
			UPSInterfaceBase UPS = new UPSInterfaceBase();
			UPS.Connect();
			bool open = UPS.USBPort.Open();
			if(open)
            {
				if (!File.Exists("BMGUPS.config"))
				{
				try
				{
					byte[] request = new byte[]
					{
						81,
						77,
						68,
						13,
						0,
						0,
						0,
						0
					};
					string[] array = UPS.USBPort.Getinfo(request).Replace("#", "").Replace("(", "").Split(new char[]{' '});
					var upsmodel = array[0] + " " + array[1] + "VA";
					UPSSize = (int.Parse(array[1]) * int.Parse(array[2]) / 100);
					var UPSBattVtxt = (double.Parse(array[6]) * double.Parse(array[7])).ToString();
					int.TryParse(array[6], out Batterys);
				}
				catch{}
					File.WriteAllText("BMGUPS.config", UPSSize + "|" + Batterys + "|" + CostKWh + "|" + PowerFactor);
				}
				else
				{
					string[] config = File.ReadAllText("BMGUPS.config").Split(new string[] { "|" }, StringSplitOptions.None);
					if (config.Length >= 4)
					{
						int.TryParse(config[0], out UPSSize);
						int.TryParse(config[1], out Batterys);
						float.TryParse(config[2], out float CostKWh);
                        float.TryParse(config[3], out float PowerFactor);
                    }
				}
			}
			byte loops = 0;
			while (open)
			{
				Thread.Sleep(1000);
				var packet = UPS.LoadData();
				ClearConsoleSmoothly();
				Console.WriteLine($"AC In: {packet.ACVoltage}v");
				Console.WriteLine($"AC Out: {packet.LoadVoltage}v");
				Console.WriteLine($"Battery: {packet.BatteryVoltage}v");
				Console.WriteLine($"Load: {packet.Load}%");
				Console.WriteLine($"Temp: {packet.Temperature}c");
				Console.WriteLine($"Watts: {packet.Watts}w");
				Console.WriteLine($"Time Remaining: {packet.TimeLeft}m");
				w += (ulong)packet.Watts;
				double cost = ((w / 3600000.0) / PowerFactor) * CostKWh ;
				Console.WriteLine($"Cost: ${cost:F5}");
				if (DumpCSV)
				{
					try
					{
						string csvLine = string.Join(",",
							packet.ACVoltage,
							packet.LoadVoltage,
							packet.BatteryVoltage,
							packet.Load,
							packet.Temperature,
							packet.Watts,
							packet.TimeLeft);
						File.AppendAllText(FileDump, csvLine + Environment.NewLine);
					}
					catch (Exception ex)
					{
						Console.WriteLine($"Failed to write dump: {ex.Message}");
					}
				}
				loops++;
				if(loops >= byte.MaxValue){SaveData(); loops = 0; }
			}
			UPS.Disconnect();
		}

	
		static void ClearConsoleSmoothly()
		{
			int width = Console.WindowWidth;
			int height = Console.WindowHeight;
			Console.SetCursorPosition(0, 0);
			string blankLine = new string(' ', width);
			for (int i = 0; i < height; i++){Console.Write(blankLine);}
			Console.SetCursorPosition(0, 0); 
		}

		static void SaveData(){File.WriteAllText("KWh.txt", w.ToString());}

		public static class Win32Hid
		{
			[StructLayout(LayoutKind.Sequential, Pack = 1)]
			public struct SP_DEVICE_INTERFACE_DATA
			{
				public int Size;
				public Guid InterfaceClassGuid;
				public int Flags;
				public IntPtr Reserved;
			}

			[StructLayout(LayoutKind.Sequential, Pack = 1)]
			public struct SP_DEVICE_INTERFACE_DETAIL_DATA
			{
				public int Size;
				[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
				public string DevicePath;
			}

			[StructLayout(LayoutKind.Sequential, Pack = 1)]
			public struct HIDP_CAPS
			{
				public ushort Usage;
				public ushort UsagePage;
				public ushort InputReportByteLength;
				public ushort OutputReportByteLength;
				public ushort FeatureReportByteLength;
				[MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
				public ushort[] Reserved;
				public ushort NumberLinkCollectionNodes;
				public ushort NumberInputButtonCaps;
				public ushort NumberInputValueCaps;
				public ushort NumberInputDataIndices;
				public ushort NumberOutputButtonCaps;
				public ushort NumberOutputValueCaps;
				public ushort NumberOutputDataIndices;
				public ushort NumberFeatureButtonCaps;
				public ushort NumberFeatureValueCaps;
				public ushort NumberFeatureDataIndices;
			}

			[StructLayout(LayoutKind.Sequential, Pack = 1)]
			public struct HIDD_ATTRIBUTES
			{
				public uint Size;
				public ushort VendorID;
				public ushort ProductID;
				public ushort VersionNumber;
			}

			public const int DIGCF_PRESENT = 2;
			public const int DIGCF_DEVICEINTERFACE = 16;
			public const uint FILE_FLAG_OVERLAPPED = 1073741824u;
			public static string[] DevicePaths
			{
				get
				{
					List<string> list = new List<string>();
					HidD_GetHidGuid(out var gHid);
					IntPtr intPtr = SetupDiGetClassDevs(ref gHid, null, IntPtr.Zero, 18u);
					try
					{
						SP_DEVICE_INTERFACE_DATA oInterfaceData = default(SP_DEVICE_INTERFACE_DATA);
						oInterfaceData.Size = Marshal.SizeOf(oInterfaceData);
						uint num = 0u;
						while (SetupDiEnumDeviceInterfaces(intPtr, 0u, ref gHid, num++, ref oInterfaceData))
						{
							SP_DEVICE_INTERFACE_DETAIL_DATA oDetailData = default(SP_DEVICE_INTERFACE_DETAIL_DATA);
							if (IntPtr.Size == 8){oDetailData.Size = 8;}
							else{oDetailData.Size = 5;}

							int nDeviceInterfaceDetailDataSize = Marshal.OffsetOf(typeof(SP_DEVICE_INTERFACE_DETAIL_DATA), "DevicePath").ToInt32() + Marshal.SystemDefaultCharSize * 256;
							if (SetupDiGetDeviceInterfaceDetail(intPtr, ref oInterfaceData, ref oDetailData, nDeviceInterfaceDetailDataSize, out var _, IntPtr.Zero))
							{
								list.Add(oDetailData.DevicePath);
							}
						}
					}
					finally
					{
						SetupDiDestroyDeviceInfoList(intPtr);
					}

					return list.ToArray();
				}
			}

			[DllImport("setupapi.dll", SetLastError = true)]
			public static extern IntPtr SetupDiGetClassDevs(ref Guid gClass, [MarshalAs(UnmanagedType.LPStr)] string strEnumerator, IntPtr hParent, uint nFlags);
			[DllImport("setupapi.dll", SetLastError = true)]
			public static extern int SetupDiDestroyDeviceInfoList(IntPtr lpInfoSet);
			[DllImport("setupapi.dll", SetLastError = true)]
			public static extern bool SetupDiEnumDeviceInterfaces(IntPtr lpDeviceInfoSet, uint nDeviceInfoData, ref Guid gClass, uint nIndex, ref SP_DEVICE_INTERFACE_DATA oInterfaceData);
			[DllImport("setupapi.dll", SetLastError = true)]
			public static extern bool SetupDiGetDeviceInterfaceDetail(IntPtr lpDeviceInfoSet, ref SP_DEVICE_INTERFACE_DATA oInterfaceData, ref SP_DEVICE_INTERFACE_DETAIL_DATA oDetailData, int nDeviceInterfaceDetailDataSize, out int nRequiredSize, IntPtr lpDeviceInfoData);
			[DllImport("kernel32.dll", SetLastError = true)]
			public static extern IntPtr CreateFile([MarshalAs(UnmanagedType.LPStr)] string strName, uint nAccess, uint nShareMode, IntPtr lpSecurity, uint nCreationFlags, uint nAttributes, IntPtr lpTemplate);
			[DllImport("kernel32.dll", SetLastError = true)]
			public static extern int CloseHandle(IntPtr hFile);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern void HidD_GetHidGuid(out Guid gHid);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern bool HidD_GetPreparsedData(IntPtr hFile, out IntPtr lpData);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern bool HidD_FreePreparsedData(ref IntPtr pData);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern int HidP_GetCaps(IntPtr lpData, out HIDP_CAPS oCaps);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern bool HidD_GetAttributes(IntPtr HidDeviceObject, ref HIDD_ATTRIBUTES Attributes);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern bool HidD_GetManufacturerString(IntPtr HidDeviceObject, byte[] Buffer, ulong BufferLength);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern bool HidD_GetProductString(IntPtr HidDeviceObject, byte[] Buffer, ulong BufferLength);
			[DllImport("hid.dll", SetLastError = true)]
			public static extern bool HidD_GetSerialNumberString(IntPtr HidDeviceObject, byte[] Buffer, ulong BufferLength);
		}

		public interface IDevice : IDisposable
		{
			ushort VendorID { get; }
			ushort ProductID { get; }
			void Write(byte ReportID, byte[] Data);
			int Read(byte ReportID, byte[] Buffer);
			byte[] WriteRead(byte ReportID, byte[] Data);
		}

		public static class DeviceFactory
		{
			public static string[] DevicePaths => Win32Hid.DevicePaths;
			public static IDevice CreateDevice(string Path)
			{
				Win32Device win32Device = new Win32Device(Path);
				Win32DeviceSet win32DeviceSet = new Win32DeviceSet(win32Device.VendorID, win32Device.ProductID);
				win32DeviceSet.AddDevice(win32Device);
				return win32DeviceSet;
			}

			private static IDevice[] ConcatenateDeviceSets(IDevice[] deviceSets)
			{
				List<Win32DeviceSet> list = new List<Win32DeviceSet>();
				foreach (IDevice device in deviceSets)
				{
					Win32DeviceSet win32DeviceSet = (Win32DeviceSet)device;
					bool flag = false;
					foreach (Win32DeviceSet item in list)
					{
						if (item.VendorID != win32DeviceSet.VendorID || item.ProductID != win32DeviceSet.ProductID){continue;}
						foreach (Win32Device unallocatedDevice in win32DeviceSet.UnallocatedDevices)
						{
							item.AddDevice(unallocatedDevice);
						}
						foreach (KeyValuePair<byte, Win32Device> device2 in win32DeviceSet.Devices)
						{
							item.AddDevice(device2.Key, device2.Value);
						}
						flag = true;
						break;
					}
					if (!flag)
					{
						list.Add(win32DeviceSet);
					}
				}
				return list.ToArray();
			}

			public static IDevice[] Enumerate()
			{
				List<IDevice> list = new List<IDevice>();
				string[] devicePaths = DevicePaths;
				foreach (string path in devicePaths)
				{
					try{list.Add(CreateDevice(path));}
					catch { }
				}
				return ConcatenateDeviceSets(list.ToArray());
			}

			public static IDevice[] Enumerate(ushort VendorID, ushort ProductID)
			{
				List<IDevice> list = new List<IDevice>();
				IDevice[] array = Enumerate();
				foreach (IDevice device in array)
				{
					if (device.VendorID == VendorID && device.ProductID == ProductID)
					{
						list.Add(device);
					}
				}

				return list.ToArray();
			}
		}

		public class Win32Device : IDisposable
		{
			private delegate object ReadDelegate(byte[] buffer);
			private delegate object WriteDelegate(byte[] data);
			protected IntPtr Handle;
			protected Win32Hid.HIDP_CAPS Capabilities;
			protected Win32Hid.HIDD_ATTRIBUTES Attributes;
			protected FileStream DataStream;
			public ushort VendorID => Attributes.VendorID;
			public ushort ProductID => Attributes.ProductID;
			public ushort Version => Attributes.VersionNumber;
			public int InputLength => Capabilities.InputReportByteLength;
			public int OutputLength => Capabilities.OutputReportByteLength;
			public object DoRead(byte[] buffer)
			{
				ReadDelegate readDelegate = DoReadHandler;
				IAsyncResult asyncResult = readDelegate.BeginInvoke(buffer, null, null);
				if (!asyncResult.IsCompleted)
				{
					asyncResult.AsyncWaitHandle.WaitOne(200, exitContext: false);
					if (!asyncResult.IsCompleted){throw new ApplicationException("Timeout");}
				}
				return readDelegate.EndInvoke((AsyncResult)asyncResult);
			}

			private object DoReadHandler(byte[] buffer)
			{
				try
				{
					return DataStream.Read(buffer, 0, buffer.Length);
				}
				catch (Exception)
				{
					return buffer.Length;
				}
			}

			public object DoWrite(byte[] data)
			{
				WriteDelegate writeDelegate = DoWriteHandler;
				IAsyncResult asyncResult = writeDelegate.BeginInvoke(data, null, null);
				if (!asyncResult.IsCompleted)
				{
					asyncResult.AsyncWaitHandle.WaitOne(100, exitContext: false);
					if (!asyncResult.IsCompleted)
					{
						throw new ApplicationException("Timeout");
					}
				}
				return writeDelegate.EndInvoke((AsyncResult)asyncResult);
			}

			private object DoWriteHandler(byte[] data)
			{
				try
				{
					DataStream.Write(data, 0, data.Length);
				}
				catch{}
				return null;
			}

			public Win32Device(string path)
			{
				Handle = GetDeviceHandle(path);
				Attributes = default(Win32Hid.HIDD_ATTRIBUTES);
				Win32Hid.HidD_GetAttributes(Handle, ref Attributes);
				if (Marshal.GetLastWin32Error() != 0)
				{
					throw new Win32Exception("Cannot get device attributes.");
				}

				Capabilities = default(Win32Hid.HIDP_CAPS);
				if (Win32Hid.HidD_GetPreparsedData(Handle, out var lpData))
				{
					try
					{
						Win32Hid.HidP_GetCaps(lpData, out Capabilities);
					}
					finally
					{
						Win32Hid.HidD_FreePreparsedData(ref lpData);
					}
				}

				if (Marshal.GetLastWin32Error() != 0)
				{
					throw new Win32Exception("Cannot get device capabilities.");
				}

				SafeFileHandle handle = new SafeFileHandle(Handle, ownsHandle: true);
				DataStream = new FileStream(handle, FileAccess.ReadWrite, Capabilities.InputReportByteLength, isAsync: true);
			}

			public void Write(byte[] data)
			{
				if (data.Length != Capabilities.OutputReportByteLength)
				{
					throw new Exception($"Data length must be {Capabilities.OutputReportByteLength} bytes.");
				}

				try
				{
					DoWrite(data);
				}
				catch (Exception)
				{
				}
			}

			public int Read(byte[] buffer)
			{
				if (buffer.Length != Capabilities.InputReportByteLength)
				{
					throw new Exception($"Buffer length must be {Capabilities.InputReportByteLength} bytes.");
				}

				try
				{
					return (int)DoRead(buffer);
				}
				catch (Exception)
				{
					return buffer.Length;
				}
			}

			protected static IntPtr GetDeviceHandle(string path)
			{
				IntPtr intPtr = Win32Hid.CreateFile(path, 3221225472u, 3u, IntPtr.Zero, 3u, 1073741824u, IntPtr.Zero);
				if (Marshal.GetLastWin32Error() != 0 || intPtr == new IntPtr(-1))
				{
					throw new Win32Exception($"Cannot create handle to device {path}");
				}

				return intPtr;
			}

			public void Dispose()
			{
				Dispose(Disposing: true);
				GC.SuppressFinalize(this);
			}

			protected virtual void Dispose(bool Disposing)
			{
				if (Disposing)
				{
					try
					{
						DataStream.Close();
					}
					catch (Exception)
					{
					}
				}

				Win32Hid.CloseHandle(Handle);
			}
		}

		public class Win32DeviceSet : IDevice, IDisposable
		{
			public readonly List<Win32Device> UnallocatedDevices;
			public readonly Dictionary<byte, Win32Device> Devices;
			private ushort _vendorID;
			private ushort _productID;
			public ushort VendorID => _vendorID;
			public ushort ProductID => _productID;
			public Win32DeviceSet(ushort VendorID, ushort ProductID)
			{
				_vendorID = VendorID;
				_productID = ProductID;
				UnallocatedDevices = new List<Win32Device>();
				Devices = new Dictionary<byte, Win32Device>();
			}

			public void AddDevice(Win32Device Device)
			{
				if (Device.VendorID != _vendorID || Device.ProductID != _productID)
				{
					throw new Exception("Device and/or Product IDs do not match this device set.");
				}
				UnallocatedDevices.Add(Device);
			}

			public void AddDevice(byte ReportID, Win32Device Device)
			{
				if (Device.VendorID != _vendorID || Device.ProductID != _productID)
				{
					throw new Exception("Device and/or Product IDs do not match this device set.");
				}

				UnallocatedDevices.Remove(Device);
				Devices.Add(ReportID, Device);
			}

			protected static byte[] AddIdToReport(byte ReportID, byte[] Data)
			{
				byte[] array = new byte[Data.Length + 1];
				array[0] = ReportID;
				Array.Copy(Data, 0, array, 1, Data.Length);
				return array;
			}

			public void Write(byte ReportID, byte[] Data)
			{
				if (!Devices.ContainsKey(ReportID))
				{
					for (int i = 0; i < UnallocatedDevices.Count; i++)
					{
						if (UnallocatedDevices[i].OutputLength == Data.Length - 1)
						{

							AddDevice(ReportID, UnallocatedDevices[i]);
							break;
						}
					}
				}

				Devices[ReportID].Write(AddIdToReport(ReportID, Data));
			}

			public int Read(byte ReportID, byte[] Buffer)
			{
				if (!Devices.ContainsKey(ReportID))
				{
					for (int i = 0; i < UnallocatedDevices.Count; i++)
					{
						if (UnallocatedDevices[i].InputLength == Buffer.Length + 1)
						{

							AddDevice(ReportID, UnallocatedDevices[i]);
							break;
						}
					}
				}

				Win32Device win32Device = Devices[ReportID];
				if (Buffer.Length + 1 != win32Device.InputLength)
				{
					throw new Exception($"Buffer length must be {win32Device.InputLength - 1} bytes.");
				}

				byte[] array = new byte[win32Device.InputLength];
				int num = Devices[ReportID].Read(array);
				Array.Copy(array, 1, Buffer, 0, num - 1);
				return num - 1;
			}

			public byte[] WriteRead(byte ReportID, byte[] Data)
			{
				Write(ReportID, Data);
				byte[] array = new byte[Devices[ReportID].InputLength - 1];
				Read(ReportID, array);
				return array;
			}

			public void Dispose()
			{
				foreach (Win32Device unallocatedDevice in UnallocatedDevices)
				{
					try
					{
						unallocatedDevice.Dispose();
					}
					catch (Exception)
					{
					}
				}

				foreach (Win32Device value in Devices.Values)
				{
					try
					{
						value.Dispose();
					}
					catch (Exception)
					{
					}
				}
			}
		}

		internal struct PacketStyleA
		{

			private string CalcTime(int Load, bool BackupOperation)
			{
				double num = 7 * (double)BatteryVoltage;
				if (BackupOperation)
				{
					if (Load == 0)
					{
						return "0";
					}
					if (0 <= Load & Load <= 6)
					{
						return Math.Round(num / 5.0, 0).ToString();
					}
					if (7 <= Load & Load <= 12)
					{
						return Math.Round(num / 10.0, 0).ToString();
					}
					if (Load >= 13)
					{
						return Math.Round(num / (double)Load, 0).ToString();
					}
				}
				else
				{
					if (Load == 0)
					{
						return "0";
					}
					if (1 <= Load & Load <= 2)
					{
						return Math.Round(num / 2.0, 0).ToString();
					}
					if (3 <= Load & Load <= 7)
					{
						return Math.Round(num / 7.0, 0).ToString();
					}
					if (Load >= 8)
					{
						return Math.Round(num / (double)Load, 0).ToString();
					}
				}
				return "";
			}

			private double CalcPower(int Load, bool BackupOperation)
			{
				if (BackupOperation)
				{
					if (Load == 0)
					{
						return 0;
					}
					if (0 <= Load & Load <= 6)
					{
						return Math.Round(UPSSize * 5.0 / 100.0, 0);
					}
					if (7 <= Load & Load <= 12)
					{
						return Math.Round(UPSSize * 10.0 / 100.0, 0);
					}
					if (Load >= 13)
					{
						return Math.Round(UPSSize * (double)Load / 100.0, 0);
					}
				}
				else
				{
					if (Load == 0)
					{
						return 0;
					}
					if (1 <= Load & Load <= 2)
					{
						return Math.Round(UPSSize * 2.0 / 100.0, 0);
					}
					if (3 <= Load & Load <= 7)
					{
						return Math.Round(UPSSize * 7.0 / 100.0, 0);
					}
					if (Load >= 8)
					{
						return Math.Round(UPSSize * (double)Load / 100.0, 0);
					}
				}
				return 0;
			}

			public void ParseString(string Input)
			{
				try
				{
					ACVoltage = float.Parse(Input.Substring(1, 3)) + float.Parse(Input.Substring(5, 1)) / 10f;
					LoadVoltage = float.Parse(Input.Substring(13, 3)) + float.Parse(Input.Substring(17, 1)) / 10f;
					Load = int.Parse(Input.Substring(19, 3));
					LoadFrequency = float.Parse(Input.Substring(23, 2)) + float.Parse(Input.Substring(26, 1)) / 10f;
					BatteryVoltage = float.Parse(Input.Substring(28, 4)) * Batterys * Batterys;
					Temperature = float.Parse(Input.Substring(33, 2)) + float.Parse(Input.Substring(36, 1)) / 10f;
					BackupOperation = (Input.Substring(38, 1) == "1");
					BatteryCritical = (Input.Substring(39, 1) == "1");
					TestRunning = (Input.Substring(43, 1) == "1");
					AudibleAlarm = (Input.Substring(45, 1) == "1");
					Watts = CalcPower(Load, BackupOperation);
					TimeLeft = CalcTime(Load, BackupOperation);
					RECEVEDCorrectly = true;
				}
				catch (Exception) { RECEVEDCorrectly = false; }
			}
			public float ACVoltage;
			public float LoadVoltage;
			public int Load;
			public double Watts;
			public string TimeLeft;
			public float LoadFrequency;
			public float BatteryVoltage;
			public float Temperature;
			public bool BackupOperation;
			public bool BatteryCritical;
			public bool TestRunning;
			public bool AudibleAlarm;
			public bool RECEVEDCorrectly;
		}

		internal struct PacketStyleB
		{
			public void ParseString(string Input)
			{
				try
				{
					ACVoltage = float.Parse(Input.Substring(1, 3)) + float.Parse(Input.Substring(5, 1)) / 10f;
					DCVoltage = float.Parse(Input.Substring(11, 2)) + float.Parse(Input.Substring(14, 2)) / 100f;
					LoadFrequency = float.Parse(Input.Substring(17, 2)) + float.Parse(Input.Substring(20, 1)) / 10f;
					RECEVEDCorrectly = true;
				}
				catch (Exception) { RECEVEDCorrectly = false; }
			}

			public float ACVoltage;
			public float LoadFrequency;
			public float DCVoltage;
			public bool RECEVEDCorrectly;
		}

		internal class UPSInterfaceBase
		{
			public UPSInterfaceBase()
			{
				USBPort = new UPSInterfaceBase.USBPortBase();
			}

			public bool Connect()
			{
				if (USBPort == null || !USBPort.Open())
				{
					return false;
				}
				if (USBPort.CallUPS())
				{
					return true;
				}
				USBPort.Close();
				return false;
			}

			public void Disconnect()
			{
				if (USBPort != null)
				{
					USBPort.Close();
				}
			}

			public PacketStyleA LoadData()
			{
				PacketStyleA packetStyleA = default(PacketStyleA);
				int num = 1;
				do
				{
					if (USBPort != null)
					{
						packetStyleA = USBPort.RequestPacketStyleA();
					}
					if (!packetStyleA.RECEVEDCorrectly)
					{
						USBPort.Close();
						Thread.Sleep(1000);
						if (!USBPort.Open())
						{
							num = 2;
						}
					}
					num++;
				}
				while (num <= 2 & !packetStyleA.RECEVEDCorrectly);
				if (packetStyleA.RECEVEDCorrectly)
				{
					return packetStyleA;
				}
				OnMessageConnectionLost();
				return packetStyleA;
			}

			public void OnMessageConnectionLost()
			{
				Disconnect();
				Console.Write("Connection to UPS lost!");
			}


			public UPSInterfaceBase.USBPortBase USBPort;

			public class USBPortBase
			{
				public bool Open()
				{
					IDevice[] array = DeviceFactory.Enumerate(1637, 20833);
					if (array.Length < 1)
					{
						return false;
					}
					BMGUPS = array[0];
					if (BMGUPS is Win32DeviceSet)
					{
						Win32DeviceSet win32DeviceSet = (Win32DeviceSet)BMGUPS;
						foreach (Win32Device win32Device in new List<Win32Device>(win32DeviceSet.UnallocatedDevices))
						{
							int outputLength = win32Device.OutputLength;
							if (outputLength == 9)
							{
								win32DeviceSet.AddDevice(0, win32Device);
							}
							else
							{
								win32Device.Dispose();
							}
						}
					}
					return true;
				}

				public bool CallUPS()
				{
					PacketStyleA packetStyleA = default(PacketStyleA);
					PacketStyleB packetStyleB = default(PacketStyleB);
					if (BMGUPS != null)
					{
						packetStyleA = RequestPacketStyleA();
						packetStyleB = RequestPacketStyleB();
					}
					return packetStyleA.RECEVEDCorrectly & packetStyleB.RECEVEDCorrectly;
				}


				public void Close()
				{
					if (BMGUPS != null)
					{
						BMGUPS = null;
					}
				}


				public PacketStyleA RequestPacketStyleA()
				{
					PacketStyleA result = default(PacketStyleA);
					if (BMGUPS != null)
					{
						byte[] array = new byte[]
						{
						81,
						49,
						13,
						0,
						0,
						0,
						0,
						0
						};
						byte[] array2 = BMGUPS.WriteRead(0, array);
						if (array2[0] == 40)
						{
							string text = "";
							for (int i = 1; i <= 6; i++)
							{
								for (int j = 0; j <= 7; j++)
								{
									text += Convert.ToChar(array2[j]).ToString();
								}
								if (i != 6)
								{
									BMGUPS.Read(0, array2);
								}
							}
							result.ParseString(text);
						}
					}
					return result;
				}


				public PacketStyleB RequestPacketStyleB()
				{
					PacketStyleB result = default(PacketStyleB);
					if (BMGUPS != null)
					{
						byte[] array = new byte[8];
						array[0] = 70;
						array[1] = 13;
						byte[] array2 = array;
						byte[] array3 = BMGUPS.WriteRead(0, array2);
						if (array3[0] == 35)
						{
							string text = "";
							for (int i = 1; i <= 3; i++)
							{
								for (int j = 0; j <= 7; j++)
								{
									text += Convert.ToChar(array3[j]).ToString();
								}
								if (i != 3)
								{
									BMGUPS.Read(0, array3);
								}
							}
							result.ParseString(text);
						}
					}
					return result;
				}

				public string Getinfo(byte[] request)
				{
					if (BMGUPS != null)
					{
						try
						{
							byte[] array = BMGUPS.WriteRead(0, request);
							string text = "";
							for (int i = 1; i <= 7; i++)
							{
								for (int j = 0; j <= 7; j++)
								{
									text += Convert.ToChar(array[j]).ToString();
								}
								if (i != 40)
								{
									BMGUPS.Read(0, array);
								}
							}
							if (text == "(NAK")
							{
								return "Unsupported Command";
							}
							return text;
						}
						catch
						{
						}
					}
					return "Bad Command";
				}

				private IDevice BMGUPS;
			}
		}
	}
}