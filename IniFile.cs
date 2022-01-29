using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Kred_calc
{
    // Создание нового INI-файла для хранения данных
    class IniFile
    {
        static string path = "";

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string section,string key,string val,string filePath);
        
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(string section, string key,string def, StringBuilder retVal,int size,string filePath);

        // Путь к INI-файлу
        static public void IniFile_PATH(string INIPath)
        {
            path = INIPath;
        }
        
        // Запись данных в INI-файл
        static public void IniWriteValue(string Section,string Key,string Value)
        {
            WritePrivateProfileString(Section,Key,Value, path);
        }
        
        // Чтение данных из INI-файла
        static public string IniReadValue(string Section,string Key)
        {
            StringBuilder temp = new(255);
            _ = GetPrivateProfileString(Section, Key, "", temp, 255, path);
            return temp.ToString();
        }
    }
}