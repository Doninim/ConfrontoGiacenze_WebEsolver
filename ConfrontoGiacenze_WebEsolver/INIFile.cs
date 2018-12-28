using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ConfrontoGiacenze_WebEsolver
{
    class INIFile
    {
        string Path;

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

        public INIFile(string inipath)
        {
            Path = inipath;
        }

        public string IniReadValue(string Key, string Section)
        {
            StringBuilder RetVal = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", RetVal, 255, Path);
            return RetVal.ToString();
        }
    }
}
