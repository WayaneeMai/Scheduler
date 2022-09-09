using System;
using System.IO;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Text;
using System.Linq;
using System.Data;

namespace TestingScheduling
{
    class Time_Caculator
    {
        public static double CaculateTimeSpan(DateTime ScheduleTime, DateTime Time2)//依據狀態，設定available time
        {
            if (ScheduleTime > Time2)
            {
                TimeSpan interval = ScheduleTime- Time2;
                return interval.TotalMinutes;
            }
            else
            {
                TimeSpan interval = Time2 - ScheduleTime;
                return interval.TotalMinutes;
            }
        }
    }
}
