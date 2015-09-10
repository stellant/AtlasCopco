using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.IO.Ports;
using System.Threading;
using Microsoft.Win32.SafeHandles;

using MicroTorque.Exceptions;

namespace MicroTorque
{
    public class MTCom
    {
        #region Constants & structures


        public readonly static IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);


        public enum Unit : int
        {
            MTU_mNm = 0,
            MTU_cNm,
            MTU_Nm,
            MTU_mN,
            MTU_N,
            MTU_kN,
            MTU_inlbf,
            MTU_lbf,
            MTU_inozf,
            MTU_gcm,
            MTU_kgm,
            MTU_ftlbf,
            MTU_ozf,
            MTU_kgf,
            MTU_gf,
            MTU_INVALID = -1,
        }


        public enum DeviceStatus : int
        {
            MT_DEVICE_NOT_FOUND = 0,
            MT_DEVICE_PRESENT,
            MT_DEVICE_CONNECTED,
            MT_DEVICE_READY
        }


        public enum DeviceType : int
        {
            MT_DEVICE_UNKNOWN = 0,
            MT_DEVICE_MICROTEST_MC,
            MT_DEVICE_MT_G4,
            MT_DEVICE_ACTA_MT4,
            MT_DEVICE_MTF400_B,
            MT_DEVICE_MTF400_A,
            MT_DEVICE_MTF400_D,
            MT_DEVICE_MTF6000,
        }


        public enum ErrorCode : int
        {
            MT_OK = 0,
            MT_ERR_INVALID_VERSION = -1,
            MT_ERR_INVALID_PARAMETER = -2,
            MT_ERR_INVALID_DATA = -3,
            MT_ERR_INVALID_CHECKSUM = -4,
            MT_ERR_CONNECTION_ERROR = -5,
            MT_ERR_INVALID_MESSAGE_ID = -6,
            MT_ERR_MUTEX_TIMEOUT = -7,
            MT_ERR_PIPE_READ_FAILED = -8,
            MT_ERR_PIPE_WRITE_FAILED = -9,
            MT_ERR_USB_READ_FAILED = -10,
            MT_ERR_USB_WRITE_FAILED = -11,
            MT_ERR_INVALID_EVENT_OBJECT = -12,
            MT_ERR_CONNECTION_STALLING = -13,
            MT_ERR_NOT_SUPPORTED = -14,
            MT_ERR_INTERFACE_ERROR = -15,
            MT_ERR_TRANSMIT_ERROR = -16,
            MT_ERR_RECEIVE_ERROR = -17,
            MT_ERR_INVALID_HANDLE = -18,
            MT_ERR_INSUFFICIENT_SPACE = -19,
            MT_ERR_DISCONNECTED = -20,
            MT_ERR_GENERAL = -21,
            MT_ERR_NO_MORE_ITEMS = -22,
        }


        [StructLayout(LayoutKind.Sequential, Pack = 1, Size = 16)]
        public struct TracePoint
        {
            [MarshalAs(UnmanagedType.R8)]
            public double torque;
            [MarshalAs(UnmanagedType.R8)]
            public double angle;
        }


        [StructLayout(LayoutKind.Sequential, Pack = 1, Size = 96, CharSet = CharSet.Ansi)]
        public struct SDeviceInfo
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 16)]
            public string serial;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string devId;
            [MarshalAs(UnmanagedType.I4)]
            public ComInterface ifType;
            [MarshalAs(UnmanagedType.I4)]
            public DeviceType devType;
            [MarshalAs(UnmanagedType.I4)]
            public DeviceStatus devStatus;
            [MarshalAs(UnmanagedType.I4)]
            public int devIndex;
        }


        public class DeviceInfo
        {
            public static readonly DeviceInfo Empty = new DeviceInfo();

            public string Serial { get; set; }
            public string DevId { get; set; }
            public ComInterface IfType { get; set; }
            public DeviceType DevType { get; set; }
            public DeviceStatus DevStatus { get; set; }
            public int DevIndex { get; set; }

            public DeviceInfo(ref SDeviceInfo di)
            {
                Serial = di.serial;
                DevId = di.devId;
                IfType = di.ifType;
                DevType = di.devType;
                DevStatus = di.devStatus;
                DevIndex = di.devIndex;
            }

            private DeviceInfo()
            {
                Serial = "";
                DevId = "";
                IfType = ComInterface.MT_IF_INVALID;
                DevType = DeviceType.MT_DEVICE_UNKNOWN;
                DevStatus = DeviceStatus.MT_DEVICE_NOT_FOUND;
                DevIndex = -1;
            }

            public override bool Equals(object obj)
            {
                DeviceInfo di = obj as DeviceInfo;
                return ((di != null) ? this.DevId.Equals(di.DevId) : false);
            }

            public override int GetHashCode()
            {
                return DevId.GetHashCode();
            }

            public override string ToString()
            {
                return Serial;
            }
        }


        public enum ComInterface : int
        {
            MT_IF_INVALID = 0,


            MT_IF_USB,


            MT_IF_RS232,

            MT_IF_TCP,
        }


        public enum ComProtocol : int
        {
            MT_PROT_INVALID = 0,

            MT_PROT_ASCII,

            MT_PROT_LEGACY_TRACE,

            MT_PROT_LEGACY_SUMMARY,
        }

        #endregion

        #region Imported DLL functions

        internal static class SafeNativeMethods
        {
            #region MTCom v1.x DLL functions

            /// <summary>
            /// Initialization function for MTCom. Must be called before any other function is used.
            /// </summary>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            /// <remarks>As of MTCom v1.9+ it is not necessary to call this function unless the application
            ///          wishes to check that the connection to MTCom is valid.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_Init")]
            internal static extern bool Init();

            /// <summary>
            /// Returns the MTCom DLL version.
            /// </summary>
            /// <example>
            /// StringBuilder dllVersion = new StringBuilder(16);
            /// MTCom.GetDllVersion(dllVersion);
            /// </example>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetDllVersion", CharSet = CharSet.Ansi)]
            internal static extern void GetDllVersion(
                [MarshalAs(UnmanagedType.LPStr)] StringBuilder dllVersion);

            /// <summary>
            /// Returns a semicolon separated list of the serial numbers of connected MicroTorque devices.
            /// </summary>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            /// <example>
            /// int noDevices;
            /// StringBuilder devices = new StringBuilder(512);
            /// 
            /// if (MTCom.GetDeviceList(devices, out noDevices))
            /// {
            ///     Array deviceArray = devices.ToString().Split(';');
            /// }
            /// </example>
            /// <remarks>Deprecated, use CreateDeviceListEx and GetDeviceInfoEx.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetDeviceList", CharSet = CharSet.Ansi)]
            internal static extern bool GetDeviceList(
                [MarshalAs(UnmanagedType.LPStr)] StringBuilder devices,
                [MarshalAs(UnmanagedType.I4)] out int numDevices);

            /// <summary>
            /// Returns status and device type of a device with given serial number.
            /// </summary>
            /// <param name="serial">Serial number of device.</param>
            /// <param name="deviceStatus">Variable to store device status in.</param>
            /// <param name="deviceType">Variable to store device type in.</param>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            /// <remarks>Deprecated, use GetDeviceInfoEx.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetDeviceInfo", CharSet = CharSet.Ansi)]
            internal static extern bool GetDeviceInfo(
                [MarshalAs(UnmanagedType.LPStr)] string serial,
                [MarshalAs(UnmanagedType.I4)] out DeviceStatus deviceStatus,
                [MarshalAs(UnmanagedType.I4)] out DeviceType deviceType);

            /// <summary>
            /// Forces a device to refresh its device info.
            /// </summary>
            /// <param name="serial">Serial number of device.</param>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_RefreshDeviceInfo", CharSet = CharSet.Ansi)]
            internal static extern bool RefreshDeviceInfo(
                [MarshalAs(UnmanagedType.LPStr)] string serial);

            /// <summary>
            /// Opens a connection to a MicroTorque device.
            /// </summary>
            /// <param name="serial">Serial number of device to connect to.</param>
            /// <param name="reserved">This parameter is reserved for future use and should be set to 0.</param>
            /// <returns>INVALID_HANDLE_VALUE on failure (refer to GetLastError).</returns>
            /// <example>
            /// IntPtr client = MTCom.Open("127CA793", 0);
            /// if (client != MTCom.INVALID_HANDLE_VALUE)
            /// {
            ///     // Success, we now have a valid connection.
            ///     MTCom.Close(ref client);
            /// }
            /// </example>
            /// <remarks>Deprecated, use OpenEx.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_Open", CharSet = CharSet.Ansi)]
            internal static extern IntPtr Open(
                [MarshalAs(UnmanagedType.LPStr)] string serial,
                int reserved);

            /// <summary>
            /// Verifies that a connection handle is valid.
            /// </summary>
            /// <param name="client">Client connection handle.</param>
            /// <returns>True if the handle is valid, false otherwise (refer to GetLastError).</returns>
            /// <remarks>If a connection to a device has been lost, GetLastError will indicate MT_ERR_DISCONNECTED.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_IsHandleValid")]
            internal static extern bool IsHandleValid(
                IntPtr client);

            /// <summary>
            /// Close the connection to a MicroTorque device.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            [DllImport("MTCom.dll", EntryPoint = "MT_Close")]
            internal static extern void Close(ref IntPtr client);

            /// <summary>
            /// Clear any existing datasets on a specified channel. If used on a connection opened with OpenEx channel is irrelevant and can be set to 0.
            /// </summary>
            /// <remarks>
            /// Can be used to clear summary data for a device if called with channel number 13 or 14.
            /// </remarks>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_Clear")]
            internal static extern bool Clear(IntPtr client, int channel);

            /// <summary>
            /// Write a dataset to a MicroTorque device.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            /// <param name="channel">Communication channel to use.</param>
            /// <param name="dataset">Byte array containing data to write.</param>
            /// <param name="noBytes">Number of bytes to write.</param>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            /// <example>
            /// // Request the version number string of a device.
            /// byte[] dataSet = System.Text.Encoding.ASCII.GetBytes("IV");
            /// if (MTCom.WriteSet(client, channelNo, dataSet, dataSet.Length))
            /// {
            ///     // Success, we should now read back the reply using MTCom.ReadSet.
            /// }
            /// else
            /// {
            ///     // Failed, refer to GetLastError for more details.
            /// }
            /// </example>
            /// <remarks>Deprecated, use WriteSetEx.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_WriteSet")]
            internal static extern bool WriteSet(IntPtr client, int channel,
                [MarshalAs(UnmanagedType.LPArray)] byte[] dataset, int noBytes);

            /// <summary>
            /// Read a dataset from a MicroTorque device.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            /// <param name="channel">Communication channel to use.</param>
            /// <param name="buffer">Byte array to store data in. Must be able to hold 4kb of data.</param>
            /// <param name="noBytes">Integer in where to store number of bytes received.</param>
            /// <param name="timeOut">Time in milliseconds that the function will wait for a dataset before returning.</param>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            /// <example>
            /// byte[] buffer = new byte[4096];
            /// int noBytes;
            /// 
            /// if (MTCom.ReadSet(client, channelNo, buffer, out noBytes))
            /// {
            ///     if (noBytes > 0)
            ///     {
            ///         // Success, buffer will hold the dataset retrieved.
            ///     }
            ///     else
            ///     {
            ///         // No data available at this time.
            ///     }
            /// }
            /// else
            /// {
            ///     // Failed, refer to GetLastError for more details.
            /// }
            /// </example>
            /// <remarks>Deprecated, use ReadSetEx.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_ReadSet")]
            internal static extern bool ReadSet(IntPtr client, int channel,
                [MarshalAs(UnmanagedType.LPArray), Out] byte[] buffer,
                [MarshalAs(UnmanagedType.I4)] out int noBytes,
                int timeOut);

            /// <summary>
            /// Retrieve the last tightening summary received by MTCom.
            /// </summary>
            /// <param name="client">Handle to an open G4/MTF400 device.</param>
            /// <param name="summary">Byte array to store summary data in, must be able to hold 1 kb of data.</param>
            /// <param name="noBytes">Variable in which to store number of bytes retrieved.</param>
            /// <param name="sequenceNo">Variable in which to store sequence number of the summary.</param>
            /// <returns>TRUE if successful, FALSE otherwise (refer to GetLastError).</returns>
            /// <remarks>
            /// Summary data is automatically sent on channel 13 and 14 by the MicroTorque device when a joint is
            /// completed. Since this function simply returns the last summary as received by MTCom, an application
            /// that relies on this function to retrieve summary data should keep track of the sequence number to
            /// make sure that the correct summary was retrieved.
            /// </remarks>
            /// <example>
            /// <!--
            /// // Initialization part of application, read out current sequence number.
            /// byte[] dummyBuf = new byte[1024];
            /// int dummyBytes, previousSequenceNo = 0;
            /// MTCom.GetSummary(client, dummyBuf, out dummyBytes, out previousSequenceNo)
            /// ...
            /// while (applicationIsRunning)
            /// {
            ///     // Tool is started and the application checks that the BUSY signal goes from HIGH to LOW
            ///     success = StartDriverAndWait(handle, CHANNEL, 5000);
            ///     
            ///     // Retrieve summary for the joint
            ///     if (success)
            ///     {
            ///         int sequenceNo, noBytes;
            ///         byte[] summary = new byte[1024];
            ///         
            ///         for (int retries = 0; retries < 10; retries++)
            ///         {
            ///             if (MTCom.GetSummary(client, summary, out noBytes, out sequenceNo))
            ///             {
            ///                 if (sequenceNo != previousSequenceNo)
            ///                 {
            ///                     // Summary received
            /// 
            ///                     // Save sequence number of received summary for next time
            ///                     previousSequenceNo = sequenceNo;
            ///                 }
            ///             }
            ///             
            ///             System.Threading.Thread.Sleep(50);
            ///         }
            ///     }
            /// }
            /// -->
            /// </example>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetSummary")]
            internal static extern bool GetSummary(IntPtr client,
                [MarshalAs(UnmanagedType.LPArray), Out] byte[] summary,
                [MarshalAs(UnmanagedType.I4)] out int noBytes,
                [MarshalAs(UnmanagedType.I4)] out int sequenceNo);

            /// <summary>
            /// Request number of trace channels available for a device.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            /// <param name="noTraceChannels">Variable to store number of trace channels in.</param>
            /// <returns>TRUE if successful, FALSE on error (refer to GetLastError).</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetNoTraceChannels")]
            internal static extern bool GetNoTraceChannels(IntPtr client,
                [MarshalAs(UnmanagedType.I4)] out int noTraceChannels);

            /// <summary>
            /// Request information of a trace channel.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            /// <param name="traceChannel">Trace channel to request information from.</param>
            /// <param name="noPoints">Variable to store number of available data points in.</param>
            /// <param name="sampleRate">Variable to store sample rate of channel in.</param>
            /// <param name="torqueUnit">Variable to store torque unit in.</param>
            /// <returns>TRUE if successful, FALSE on error (refer to GetLastError).</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetTraceInfo")]
            internal static extern bool GetTraceInfo(IntPtr client, int traceChannel,
                [MarshalAs(UnmanagedType.I4)] out int noPoints,
                [MarshalAs(UnmanagedType.I4)] out int sampleRate,
                [MarshalAs(UnmanagedType.I4)] out int torqueUnit);

            /// <summary>
            /// Request data points from a trace channel.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            /// <param name="traceChannel">Trace channel to request data from.</param>
            /// <param name="points">Array to store data points in.</param>
            /// <param name="startPoint">Index of data point to start with.</param>
            /// <param name="maxPoints">Max number of data points to retrieve.</param>
            /// <param name="noPoints">Variable to store number of points retrieved in.</param>
            /// <returns>TRUE if successful, FALSE on error (refer to GetLastError).</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetTracePoints")]
            internal static extern bool GetTracePoints(IntPtr client, int traceChannel,
                [MarshalAs(UnmanagedType.LPArray), Out] TracePoint[] points, int startPoint, int maxPoints,
                [MarshalAs(UnmanagedType.I4)] out int noPoints);

            /// <summary>
            /// Request status of internal outputs of device.
            /// </summary>
            /// <param name="client">Handle to an open MicroTorque device.</param>
            /// <param name="output">Variable to store state of internal outputs in.</param>
            /// <returns>TRUE if successful, FALSE on error (refer to GetLastError).</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetOutput")]
            internal static extern bool GetOutput(IntPtr client,
                [MarshalAs(UnmanagedType.I4)] out int output);

            /// <summary>
            /// Request detailed error code of last called function.
            /// </summary>
            /// <returns>Error code of last called function.</returns>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetLastError")]
            internal static extern ErrorCode GetLastError();

            #endregion

            #region Imported DLL functions (v2.0+)

            /// <summary>
            /// Opens a connection to a MicroTorque device.
            /// </summary>
            /// <param name="comIf">Communication interface.</param>
            /// <param name="connStr">Connection string, depends on communication interface (see <see cref="ComInterface"/> enum for more details).</param>
            /// <param name="protocol">Communication protocol to use.</param>
            /// <param name="options">Protocol specific options. Reserved for future use, set to zero.</param>
            /// <returns>INVALID_HANDLE_VALUE on failure (refer to GetLastError).</returns>
            /// <remarks>Requires MTCom v2.0+</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_OpenEx", CharSet = CharSet.Ansi)]
            public static extern IntPtr OpenEx(
                [MarshalAs(UnmanagedType.I4)] ComInterface comIf,
                [MarshalAs(UnmanagedType.LPStr)] string connStr,
                [MarshalAs(UnmanagedType.I4)] ComProtocol protocol,
                [MarshalAs(UnmanagedType.I4)] int options);

            /// <summary>
            /// Send data to a MicroTorque device.
            /// </summary>
            /// <param name="client">Opened connection handle.</param>
            /// <param name="txBuf">Buffer containing data to send.</param>
            /// <param name="byteCount">Number of bytes to send.</param>
            /// <returns>true on success, false on failure. If the function fails and 
            ///          GetLastError returns MT_ERR_INSUFFICIENT_SPACE, then dataSize 
            ///          is set to the required size of rxBuf for the function to succeed.</returns>
            /// <remarks>Requires MTCom v2.0+</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_WriteSetEx")]
            public static extern bool WriteSetEx(IntPtr client,
                [MarshalAs(UnmanagedType.LPArray)] byte[] txBuf, int byteCount);

            /// <summary>
            /// Receive data from a MicroTorque device.
            /// </summary>
            /// <param name="client">Opened connection handle.</param>
            /// <param name="rxBuf">Buffer to store data in.</param>
            /// <param name="rxBufSize">Size of read buffer (suggested size is 4096 bytes).</param>
            /// <param name="bytesReceived">Number of bytes written to rxBuf.</param>
            /// <param name="timeOut"></param>
            /// <returns>The function returns true on success or timeout (in which case bytesReceived = 0).
            ///          If the function fails it returns false and the application can call GetLastError
            ///          for a more specific error code.</returns>
            /// <remarks>Requires MTCom v2.0+</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_ReadSetEx")]
            public static extern bool ReadSetEx(IntPtr client,
                [MarshalAs(UnmanagedType.LPArray), Out] byte[] rxBuf,
                [MarshalAs(UnmanagedType.I4)] int rxBufSize,
                [MarshalAs(UnmanagedType.I4)] out int bytesReceived,
                [MarshalAs(UnmanagedType.U4)] int timeOut);

            /// <summary>
            /// Returns the global event name for a specified connection.
            /// </summary>
            /// <param name="client">Opened connection handle.</param>
            /// <param name="nameBuf">StringBuilder object to store event name in.</param>
            /// <param name="nameBufSize">Capacity of the StringBuilder object.</param>
            /// <param name="size">Number of bytes copied to nameBuf (including NULL terminator).</param>
            /// <returns>true on success, false on failure. If the function fails and 
            ///          GetLastError returns MT_ERR_INSUFFICIENT_SPACE then size is set to
            ///          the required size of nameBuf for the function to succeed.</returns>
            /// <remarks>Requires MTCom v2.0+</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetRxEventName", CharSet = CharSet.Ansi)]
            public static extern bool GetRxEventName(IntPtr client,
                [MarshalAs(UnmanagedType.LPStr)] StringBuilder nameBuf,
                [MarshalAs(UnmanagedType.I4)] int nameBufSize,
                [MarshalAs(UnmanagedType.I4)] out int size);

            /// <summary>
            /// Creates an internal list of connected devices, information of each device retrieved using MT_GetDeviceInfoEx. 
            /// The internal device list is kept intact for each calling thread until the function is called again by the same thread.
            /// </summary>
            /// <param name="deviceCount">Number of connected devices.</param>
            /// <returns>True on success, false on failure.</returns>
            /// <remarks>Requires MTCom v2.0+</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_CreateDeviceListEx")]
            public static extern bool CreateDeviceListEx(out int deviceCount);

            /// <summary>
            /// Retrieve information of a connected device from the internal device list.
            /// The internal device list only contains connected devices, use QueryDeviceInfoEx to
            /// retrieve information of not yet connected devices.
            /// </summary>
            /// <param name="index">Index of device (0 .. (deviceCount - 1)).</param>
            /// <param name="devInfo">Structure to store device information in.</param>
            /// <param name="devInfoSize">Size of devInfo.</param>
            /// <returns></returns>
            /// <remarks>Requires MTCom v2.0+</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_GetDeviceEntryEx")]
            public static extern bool GetDeviceEntryEx(int index,
               out SDeviceInfo devInfo, int devInfoSize);

            /// <summary>
            /// Attempt to query a device for device information.
            /// </summary>
            /// <param name="connStr">Connection string to identify device to query.</param>
            /// <param name="devInfo">Structure to store device information in.</param>
            /// <param name="devInfoSize">Size of devInfo.</param>
            /// <returns>True on success, false on failure.</returns>
            /// <remarks>If a device could not be found or identified the function will return
            ///          true and devInfo.devStatus will be set to MT_DEVICE_NOT_FOUND.</remarks>
            [DllImport("MTCom.dll", EntryPoint = "MT_QueryDeviceInfoEx", CharSet = CharSet.Ansi)]
            public static extern bool QueryDeviceInfoEx(
                [MarshalAs(UnmanagedType.LPStr)] string connStr,
                out SDeviceInfo devInfo,
                int devInfoSize);
            #endregion

            #region Win32 API functions
            /// <summary>
            /// Win32 API OpenEvent.
            /// </summary>
            [DllImport("kernel32.dll", EntryPoint = "OpenEventA", CharSet = CharSet.Ansi, SetLastError = true)]
            internal static extern SafeWaitHandle OpenEvent(uint dwDesiredAccess, bool bInheritHandle,
                [MarshalAs(UnmanagedType.LPStr)] string lpName);
            #endregion
        }

        #endregion

        #region Utility functions

        /// <summary>
        /// Retrieves a manual reset event that is set when data is available on specified connection.
        /// </summary>
        /// <param name="client">Opened connection handle.</param>
        /// <returns>Event object or null on failure.</returns>
        public static ManualResetEvent GetRxEvent(IntPtr client)
        {
            StringBuilder eventName = new StringBuilder(50);
            int size;

            if (SafeNativeMethods.GetRxEventName(client, eventName, eventName.Capacity, out size))
            {
                const uint SYNCHRONIZE = 0x00100000;
                SafeWaitHandle evtHandle = SafeNativeMethods.OpenEvent(SYNCHRONIZE, true, eventName.ToString());

                if (!evtHandle.IsInvalid)
                {
                    ManualResetEvent evt = new ManualResetEvent(false);
                    evt.SafeWaitHandle = evtHandle;
                    return evt;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns a list of MicroTorque devices connected to this PC.
        /// </summary>
        /// <returns>List of MicroTorque devices.</returns>
        /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
        public static MTCom.DeviceInfo[] GetDeviceList()
        {
            int deviceCount = 0;
            MTCom.DeviceInfo[] devInfoList = new DeviceInfo[0];

            if (SafeNativeMethods.CreateDeviceListEx(out deviceCount))
            {
                if (deviceCount > 0)
                {
                    MTCom.SDeviceInfo devInfo = new MTCom.SDeviceInfo();
                    devInfoList = new MTCom.DeviceInfo[deviceCount];

                    for (int i = 0; i < deviceCount; i++)
                    {
                        if (SafeNativeMethods.GetDeviceEntryEx(i, out devInfo, Marshal.SizeOf(devInfo)))
                        {
                            devInfoList[i] = new DeviceInfo(ref devInfo);
                        }
                        else
                        {
                            devInfoList = null;
                            deviceCount = 0;
                            break;
                        }
                    }
                }
            }
            else
            {
                throw new MTComException("CreateDeviceListEx", MTCom.SafeNativeMethods.GetLastError());
            }

            return devInfoList;
        }
        #endregion


        public class Connection : IDisposable
        {
            #region Members
            IntPtr mtHandle = MTCom.INVALID_HANDLE_VALUE;
            bool isDisposed = false;
            ManualResetEvent rxEvent = null;
            ManualResetEvent allowAsyncReadEvent = null;
            Thread asyncThread = null;
            volatile bool exitAsyncThread = false;
            #endregion

            #region Properties
            public MTCom.DeviceInfo DevInfo { get; private set; }
            public MTCom.ComProtocol Protocol { get; private set; }
            public IntPtr Handle { get { return mtHandle; } }
            public bool IsOpen { get; private set; }
            #endregion

            #region Events
            public class DataEventArgs : EventArgs
            {
                public byte[] Data { get; private set; }

                public DataEventArgs(byte[] data, int length)
                {
                    Data = new byte[length];
                    Array.Copy(data, Data, length);
                }
            }

            /// <summary>
            /// Called when the asynchronous thread has received data.
            /// NOTE: Called in context of asynhronous read thread.
            /// </summary>
            public event EventHandler<DataEventArgs> OnAsyncData;

            /// <summary>
            /// Called when the connection was unexpectedly closed.
            /// NOTE: This may be called in context of asynchronous read thread.
            /// </summary>
            public event EventHandler OnAbruptDisconnect;

            /// <summary>
            /// Called when a message has been sent to MTCom.
            /// </summary>
            public event EventHandler<DataEventArgs> OnDataSent;

            /// <summary>
            /// Called when a message has been received from MTCom.
            /// </summary>
            public event EventHandler<DataEventArgs> OnDataReceived;
            #endregion

            #region Constructors and finalizers
            public Connection()
            {
                Protocol = MTCom.ComProtocol.MT_PROT_INVALID;
                DevInfo = DeviceInfo.Empty;
                IsOpen = false;
            }

            ~Connection()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!isDisposed)
                {
                    StopAsyncThread();

                    if (disposing)
                    {
                        if (rxEvent != null)
                        {
                            rxEvent.Close();
                            rxEvent = null;
                        }

                        if (allowAsyncReadEvent != null)
                        {
                            allowAsyncReadEvent.Close();
                            allowAsyncReadEvent = null;
                        }
                    }

                    if (mtHandle != MTCom.INVALID_HANDLE_VALUE)
                    {
                        MTCom.SafeNativeMethods.Close(ref mtHandle);
                    }

                    if (disposing)
                    {
                        GC.SuppressFinalize(this);
                    }

                    IsOpen = false;
                    isDisposed = true;
                }
            }
            #endregion

            #region Public methods
            /// <summary>
            /// Opens a connection to a MicroTorque device.
            /// </summary>
            /// <param name="connStr">Connection string of device to open, see documentation of MTCom.ComInterface.</param>
            /// <param name="protocol">Protocol to use.</param>
            /// <param name="param">Protocol parameter (reserved for future use).</param>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public virtual void Open(string connStr, MTCom.ComProtocol protocol, Int32 param)
            {
                MTCom.SDeviceInfo sDevInfo = new MTCom.SDeviceInfo();

                if (mtHandle != MTCom.INVALID_HANDLE_VALUE)
                    Close();

                if (!MTCom.SafeNativeMethods.QueryDeviceInfoEx(connStr, out sDevInfo, Marshal.SizeOf(sDevInfo)))
                    throw new MTComException("MTCom.QueryDeviceInfoEx", MTCom.SafeNativeMethods.GetLastError());

                if (sDevInfo.devStatus == MTCom.DeviceStatus.MT_DEVICE_READY)
                {
                    mtHandle = MTCom.SafeNativeMethods.OpenEx(sDevInfo.ifType, connStr, protocol, param);

                    if (mtHandle != MTCom.INVALID_HANDLE_VALUE)
                    {
                        rxEvent = MTCom.GetRxEvent(mtHandle);

                        if (rxEvent == null)
                        {
                            MTCom.SafeNativeMethods.Close(ref mtHandle);
                            throw new MTComException("MTCom.GetRxEvent", MTCom.SafeNativeMethods.GetLastError());
                        }

                        DevInfo = new DeviceInfo(ref sDevInfo);
                        Protocol = protocol;
                        IsOpen = true;
                    }
                    else
                    {
                        throw new MTComException("MTCom.OpenEx", MTCom.SafeNativeMethods.GetLastError());
                    }
                }
                else
                {
                    throw new MTComException(String.Format(
                        "MTCom could not find or identify device with connStr '{0}'.", connStr));
                }
            }

            /// <summary>
            /// Close an open connection.
            /// </summary>
            public virtual void Close()
            {
                StopAsyncThread();

                if (rxEvent != null)
                {
                    rxEvent.Close();
                    rxEvent = null;
                }

                if (allowAsyncReadEvent != null)
                {
                    allowAsyncReadEvent.Close();
                    allowAsyncReadEvent = null;
                }

                if (mtHandle != MTCom.INVALID_HANDLE_VALUE)
                    MTCom.SafeNativeMethods.Close(ref mtHandle);

                Protocol = MTCom.ComProtocol.MT_PROT_INVALID;
                DevInfo = DeviceInfo.Empty;
                IsOpen = false;
            }

            /// <summary>
            /// Clear any messages in MTCom receive queue for this connection.
            /// </summary>
            public void Clear()
            {
                MTCom.SafeNativeMethods.Clear(mtHandle, 0);
            }

            /// <summary>
            /// Prevents the asynchronous read thread from receiving data.
            /// </summary>
            public void DisableAsyncRead()
            {
                if (allowAsyncReadEvent != null)
                    allowAsyncReadEvent.Reset();
            }

            /// <summary>
            /// Enables the asynchronous read thread to receive data.
            /// </summary>
            public void EnableAsyncRead()
            {
                if (asyncThread == null)
                {
                    allowAsyncReadEvent = new ManualResetEvent(false);
                    exitAsyncThread = false;
                    asyncThread = new Thread(new ThreadStart(AsyncReadThreadMain));
                    asyncThread.Start();
                }

                if (allowAsyncReadEvent != null)
                    allowAsyncReadEvent.Set();
            }

            /// <summary>
            /// Send data to connected device.
            /// </summary>
            /// <param name="txData">Data buffer to send.</param>
            /// <param name="noBytes">Number of bytes in txData to send.</param>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public void Send(byte[] txData, int noBytes)
            {
                if (!MTCom.SafeNativeMethods.WriteSetEx(mtHandle, txData, noBytes))
                    HandleMTComError("MTCom.WriteSetEx");

                if (OnDataSent != null)
                    OnDataSent(this, new DataEventArgs(txData, noBytes));
            }

            /// <summary>
            /// Receive data from connected device.
            /// </summary>
            /// <param name="timeout">Timeout in milliseconds.</param>
            /// <returns>Byte array of received data. On timeout the function will return an empty array.</returns>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public byte[] Receive(int timeout)
            {
                int rxCount;
                byte[] rxData = new byte[4096];

                if (!MTCom.SafeNativeMethods.ReadSetEx(mtHandle, rxData, rxData.Length, out rxCount, timeout))
                    HandleMTComError("MTCom.ReadSetEx");

                Array.Resize(ref rxData, rxCount);

                if (OnDataReceived != null)
                    OnDataReceived(this, new DataEventArgs(rxData, rxCount));

                return rxData;
            }

            /// <summary>
            /// Read the internal output states of connected device.
            /// </summary>
            /// <returns>Device dependent 32-bit state value.</returns>
            /// <remarks>RS232 trace output must be enabled if using RS232.</remarks>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public UInt32 GetInternalOutputs()
            {
                int states;

                if (!MTCom.SafeNativeMethods.GetOutput(Handle, out states))
                    HandleMTComError("MTCom.GetOutput");

                return (UInt32)states;
            }

            /// <summary>
            /// Reads trace data received from last performed tightening or measurement.
            /// </summary>
            /// <param name="traceIndex">Trace channel index (0).</param>
            /// <param name="sampleRate">Sample rate of received trace data points.</param>
            /// <param name="torqueUnit">Torque unit of received trace data points.</param>
            /// <returns>Received trace data points (or empty array if specified trace channel doesn't exist).</returns>
            /// <remarks>RS232 trace output must be enabled if using RS232.</remarks>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            public MTCom.TracePoint[] GetTracePoints(int traceIndex, out int sampleRate, out MTCom.Unit torqueUnit)
            {
                MTCom.TracePoint[] tracePoints;
                int noPoints, recvPoints, unit;

                if (!MTCom.SafeNativeMethods.GetTraceInfo(Handle, traceIndex, out noPoints, out sampleRate, out unit))
                    HandleMTComError("MTCom.GetTraceInfo");

                tracePoints = new MTCom.TracePoint[noPoints];

                if (!MTCom.SafeNativeMethods.GetTracePoints(Handle, traceIndex, tracePoints, 0, noPoints, out recvPoints))
                    HandleMTComError("MTCom.GetTracePoints");

                if (recvPoints < tracePoints.Length)
                    Array.Resize(ref tracePoints, recvPoints);

                torqueUnit = (MTCom.Unit)unit;
                return tracePoints;
            }
            #endregion

            #region Protected methods
            /// <summary>
            /// Checks for device disconnect and throws MTCom exception on error.
            /// </summary>
            /// <param name="function"></param>
            /// <exception cref="MicroTorque.Exceptions.MTComException"></exception>
            protected void HandleMTComError(string function)
            {
                MTCom.ErrorCode mtErr = MTCom.SafeNativeMethods.GetLastError();

                if (!MTCom.SafeNativeMethods.IsHandleValid(Handle))
                    mtErr = MTCom.SafeNativeMethods.GetLastError();

                if (mtErr == ErrorCode.MT_ERR_DISCONNECTED)
                    AbruptDisconnect();

                throw new MTComException(function, mtErr);
            }
            #endregion

            #region Private methods
            /// <summary>
            /// Triggers the abrupt disconnect event.
            /// </summary>
            internal void AbruptDisconnect()
            {
                if (OnAbruptDisconnect != null)
                    OnAbruptDisconnect(this, EventArgs.Empty);
            }

            /// <summary>
            /// Asynchronous read thread, only active if asynchronous reads has been enabled.
            /// </summary>
            private void AsyncReadThreadMain()
            {
                byte[] rxBuf = new byte[4096];
                int rxSize;

                WaitHandle[] waitHandles = new WaitHandle[2];
                waitHandles[0] = rxEvent;
                waitHandles[1] = allowAsyncReadEvent;

                try
                {
                    while (!exitAsyncThread)
                    {
                        if (WaitHandle.WaitAll(waitHandles, 10))
                        {
                            if (MTCom.SafeNativeMethods.ReadSetEx(mtHandle, rxBuf, rxBuf.Length, out rxSize, 0))
                            {
                                if (OnAsyncData != null && rxSize > 0)
                                {
                                    OnAsyncData(this, new DataEventArgs(rxBuf, rxSize));
                                }
                            }
                            else
                            {
                                // An error occurred, disconnect & abort
                                AbruptDisconnect();
                                break;
                            }
                        }
                    }
                }
                catch (ThreadInterruptedException)
                {
                    // Thread interrupted, abort
                }
            }

            /// <summary>
            /// Stops the asynchronous read thread.
            /// </summary>
            private void StopAsyncThread()
            {
                Thread thread = Interlocked.Exchange(ref asyncThread, null);

                if (thread != null)
                {
                    exitAsyncThread = true;

                    if (!Thread.CurrentThread.Equals(thread))
                        thread.Join();
                }
            }
            #endregion
        }
    }

    #region MTCom exceptions
    namespace Exceptions
    {
        /// <summary>
        /// MicroTorque base exception class, all MicroTorque exceptions derive from this.
        /// </summary>
        public class MTException : Exception
        {
            public MTException(string message) : base(message) { }
        }

        /// <summary>
        /// MTCom error exception.
        /// </summary>
        public class MTComException : MTException
        {
            public MTCom.ErrorCode ErrorCode { get; private set; }

            public MTComException(string message)
                : base(message)
            {
                ErrorCode = MTCom.ErrorCode.MT_ERR_GENERAL;
            }

            public MTComException(string func, MTCom.ErrorCode code)
                : base(String.Format("{0} failed w/err: {1}", func, code))
            {
                ErrorCode = code;
            }
        }
    }
    #endregion
}
