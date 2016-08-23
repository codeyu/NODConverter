using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using unoidl.com.sun.star.ucb;

namespace NODConverter
{
    public class SocketUtils
    {
        [StructLayout(LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Ansi)]
        public struct WSAData
        {
            public short wVersion;
            public short wHighVersion;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 257)]
            public string szDescription;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 129)]
            public string szSystemStatus;
            public short iMaxSockets;
            public short iMaxUdpDg;
            public IntPtr lpVendorInfo;
        }

        [DllImport("ws2_32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        static extern Int32 WSAStartup(Int16 wVersionRequested, ref WSAData wsaData);

        [DllImport("ws2_32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
        static extern Int32 WSACleanup();


        public static void Connect()
        {
            WSAData data = new WSAData();
            int result = 0;

            data.wHighVersion = 2;
            data.wVersion = 2;

            result = WSAStartup(36, ref data);
            if (result != 0)  
            {
                throw new Exception("socket connect failure!");
            }
        }
        public static void Disconnect()
        {
            WSACleanup();
        }
    }
}
