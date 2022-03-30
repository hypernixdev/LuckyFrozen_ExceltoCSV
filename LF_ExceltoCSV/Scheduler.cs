using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using GlobalClass;

namespace LF_ExceltoCSV
{
    class Scheduler
    {
        private readonly System.Timers.Timer _timer;
        bool _isFinish = true;
        public Scheduler()
        {
            _timer = new System.Timers.Timer();
            _timer.Interval = 30000;
            _timer.Elapsed += TimerElapsed;
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            if (!_isFinish)
            {
                return;
            }

            try
            {
                _isFinish = false;
                //Thread.Sleep(30000);
                ConverttoCSV conv = new ConverttoCSV();
                conv.ConvertExcelToCSV();
                
            }
            catch (Exception ex)
            {
                Log.WriteErrorNlog("Error Converting excel to CSV. Message: " + ex.Message.ToString());
            }
            finally
            {
                _timer.Interval = 120000;
                _timer.Start();
                _isFinish = true;
            }
            
        }

        public void Start()
        {
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Stop();
        }
    }
}
