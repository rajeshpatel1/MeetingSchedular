using System;
using System.Configuration;
using System.IO;

namespace MeetingSchedularApp
{
    public class Helper
    {
        private string logFileName;

        internal Helper()
        {
            this.logFileName = string.Concat(new string[]
            {
                GetLogFilePath(),
                "MeetingSched_",
                DateTime.Now.ToString("yyyyMMdd_HHmmss"),
                ".log"
            });
        }

        internal static string GetConfigItem(string itemName)
        {
            return ConfigurationManager.AppSettings[itemName];
        }

        internal void UpdateLog(string lineText)
        {
            this.UpdateLog(lineText, true, true);
        }

        internal void UpdateLog(string lineText, bool consoleOutput, bool logFileOutput)
        {
            if (consoleOutput)
            {
                Console.WriteLine(lineText);
            }
            if (logFileOutput)
            {
                if (!System.IO.File.Exists(this.logFileName))
                {
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(this.logFileName));
                }
                var num = 0;
                while (true)
                {
                    try
                    {
                        using (TextWriter textWriter = new StreamWriter(this.logFileName, true))
                        {
                            textWriter.WriteLine(DateTime.Now.ToString(System.Globalization.CultureInfo.InvariantCulture) + " :   " + lineText);
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (num > 10)
                        {
                            Console.WriteLine("Failed to write to Log File: " + ex.Message);
                            Console.WriteLine(lineText);
                            break;
                        }
                        num++;
                        System.Threading.Thread.Sleep(1000);
                    }
                }
            }
        }

        private static string GetLogFilePath()
        {
            var text = string.Empty;
            text = GetConfigItem("LogFilePath");
            text = ((text == string.Empty) ? ".\\" : text);
            return text.EndsWith("\\", StringComparison.InvariantCulture) ? text : (text + "\\");
        }
    }
}