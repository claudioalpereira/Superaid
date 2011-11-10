using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace test
{
    class MyMail
    {
        public static void test1()
        {
            string sendEmailPath = MyConfig.SendEmailPath;
            Dictionary<string, string> env = new Dictionary<string, string>();
            env.Add("Path", sendEmailPath);
            OutRun(sendEmailPath + "test.bat", "", env);
        }
        public static void OutRun(string app, string args = "", IEnumerable<KeyValuePair<string, string>> environmentVars = null)
        {
            var pi = new ProcessStartInfo { FileName = app, Arguments = args };
            if (environmentVars != null)
                foreach (var pair in environmentVars)
                {
                    pi.UseShellExecute = false;
                    if (pi.EnvironmentVariables.ContainsKey(pair.Key))
                        pi.EnvironmentVariables[pair.Key] += ";" + pair.Value;
                    else
                        pi.EnvironmentVariables.Add(pair.Key, pair.Value);
                }
            Process.Start(pi);
        }
    }
}
