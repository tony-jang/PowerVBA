using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Core.Connector
{
    public static class VersionSelector
    {
        private const string RegKey = @"Software\Microsoft\Windows\CurrentVersion\App Paths";
        public static PPTVersion GetPPTVersion()
        {
            string path = GetPowerPointPath();
            if (File.Exists(path))
            {
                try
                {
                    FileVersionInfo FileVersion = FileVersionInfo.GetVersionInfo(path);
                    return (PPTVersion)FileVersion.FileMajorPart;
                }
                catch {}
            }

            return PPTVersion.Unknown;
        }

        public static string GetPowerPointPath()
        {
            RegistryKey key = Registry.CurrentUser;
            try
            {
                key = key.OpenSubKey(RegKey + "\\powerpnt.exe", false);
                if (key != null)
                {
                    return key.GetValue(string.Empty).ToString();
                }
            }
            catch { }

            key = Registry.LocalMachine;

            try
            {
                key = key.OpenSubKey(RegKey + "\\powerpnt.exe", false);
                if (key != null)
                {
                    return key.GetValue(string.Empty).ToString();
                }
            }
            catch { }

            return "";
        }

        
        
    }
}
