using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GlobalClass
{
    public static class Log
    {
        private static NLog.Logger loggerErr = LogManager.LoadConfiguration("NlogApp.config").GetCurrentClassLogger();
        private static bool enableInfoLog = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableInfoLog"]);
        private static bool enableInfoTimeLog = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableInfoTimeLog"]);

        public static void WriteErrorNlog(string msg)
        {
            loggerErr.Error("Error->" + msg);
        }

        public static void WriteInfoNlog(string msg)
        {
            if (enableInfoLog)
                loggerErr.Info("Info->" + msg);
        }

        public static void WriteInfoTimeNlog(string msg)
        {
            if (enableInfoTimeLog)
                loggerErr.Info("Info->" + msg);
        }
    }
}
