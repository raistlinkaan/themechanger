using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace themechanger
{
    public class ThemeFolder
    {
        private readonly string _themeFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\AppData\Local\Microsoft\Windows\Themes";

        public List<string> GetThemes()
        {
            var files = new DirectoryInfo(_themeFolder).GetFiles("*.theme", SearchOption.AllDirectories);
            return files.Select(item => item.FullName).ToList();
        }
    }
}
