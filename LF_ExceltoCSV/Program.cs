using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace LF_ExceltoCSV
{
    class Program
    {
        static void Main(string[] args)
        {

            var exitCode = HostFactory.Run(x =>
            {
                x.Service<Scheduler>(s =>
                {
                    s.ConstructUsing(schd => new Scheduler());
                    s.WhenStarted(schd => schd.Start());
                    s.WhenStopped(schd => schd.Stop());

                });

                x.RunAsLocalSystem();

                x.SetServiceName("ConvertExcelToCSV");
                x.SetDisplayName("ConvertExcelToCSV");
                x.SetDescription("Scheduler to convert Excel to CSV");
            });

            int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode());
            Environment.ExitCode = exitCodeValue;
            
        }
    }
}
