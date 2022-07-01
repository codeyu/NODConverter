using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
namespace NODConverter
{
    public static class EnvUtils
    {
        public static void InitUno()
        {
            var info = Get();
            if(info.Kind == OfficeKind.LibreOffice)
            {
                Environment.SetEnvironmentVariable("UNO_PATH", info.OfficeUnoPath, EnvironmentVariableTarget.Process);
                Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + @";" + info.OfficeUnoPath, EnvironmentVariableTarget.Process);
            }
            
        }
        public static OfficeInfo Get()
        {
            String unoPath = "";
            var libreOfficeSubKeyName = @"SOFTWARE\LibreOffice\UNO\InstallPath";
            var openOfficeSubKeyName = @"SOFTWARE\OpenOffice\UNO\InstallPath";

            // access 32bit registry entry for latest LibreOffice for Current User
            Microsoft.Win32.RegistryKey hkcuView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32);
            Microsoft.Win32.RegistryKey hkcuUnoInstallPathKey = hkcuView32.OpenSubKey(libreOfficeSubKeyName, false);

            // access 32bit registry entry for latest LibreOffice for Local Machine (All Users)
            Microsoft.Win32.RegistryKey hklmView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32);
            Microsoft.Win32.RegistryKey hklmUnoInstallPathKey = hklmView32.OpenSubKey(libreOfficeSubKeyName, false);
            if (hkcuUnoInstallPathKey != null && hkcuUnoInstallPathKey.ValueCount > 0)
            {
                unoPath = (string)hkcuUnoInstallPathKey.GetValue(hkcuUnoInstallPathKey.GetValueNames()[hkcuUnoInstallPathKey.ValueCount - 1]);
            }
            else
            {

                if (hklmUnoInstallPathKey != null && hklmUnoInstallPathKey.ValueCount > 0)
                {
                    unoPath = (string)hklmUnoInstallPathKey.GetValue(hklmUnoInstallPathKey.GetValueNames()[hklmUnoInstallPathKey.ValueCount - 1]);
                }
            }
            if (string.IsNullOrEmpty(unoPath))
            {
                hkcuUnoInstallPathKey = hkcuView32.OpenSubKey(openOfficeSubKeyName, false);
                hklmUnoInstallPathKey = hklmView32.OpenSubKey(openOfficeSubKeyName, false);
                if (hkcuUnoInstallPathKey != null && hkcuUnoInstallPathKey.ValueCount > 0)
                {
                    unoPath = (string)hkcuUnoInstallPathKey.GetValue(hkcuUnoInstallPathKey.GetValueNames()[hkcuUnoInstallPathKey.ValueCount - 1]);
                }
                else
                {

                    if (hklmUnoInstallPathKey != null && hklmUnoInstallPathKey.ValueCount > 0)
                    {
                        unoPath = (string)hklmUnoInstallPathKey.GetValue(hklmUnoInstallPathKey.GetValueNames()[hklmUnoInstallPathKey.ValueCount - 1]);
                    }
                }
            }
            
            var officeInfo = new OfficeInfo
            {
                OfficeUnoPath = unoPath,
                Kind = string.IsNullOrEmpty(unoPath) ? OfficeKind.Unknown : (unoPath.Contains("LibreOffice") ? OfficeKind.LibreOffice : OfficeKind.OpenOffice)
            };
            return officeInfo;
        }

        public static bool RunCmd(string workingPath, string cmd, string cmdArguments)
        {
            var processStartInfo = new ProcessStartInfo(cmd, cmdArguments)
				{
					WindowStyle = ProcessWindowStyle.Hidden,
					CreateNoWindow = true,
					UseShellExecute = false,
					WorkingDirectory = Path.GetDirectoryName(workingPath),
					RedirectStandardInput = true,
					RedirectStandardOutput = true,
					RedirectStandardError = true
				};

            Process p = Process.Start(processStartInfo);

            if (p.WaitForExit(20000))
            {
                return p.ExitCode == 0;
            }
            return p.ExitCode == 0;
        }
    }
    public class OfficeInfo
    {
        public string OfficeUnoPath { get; set; }
        public OfficeKind Kind { get; set; }
    }
    public enum OfficeKind
    {
        LibreOffice = 0,
        OpenOffice = 1,
        Unknown = 2,
    }

}
