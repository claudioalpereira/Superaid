using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace test
{
    class MyConfig
    {
        public static string SendEmailPath
        {
            get{ return Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["sendEmailPath"];}
            set { ConfigurationManager.AppSettings["sendEmailPath"] = value; }
        }
    }
}
