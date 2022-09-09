using System;
using System.Collections.Generic;
using System.Text;

namespace TestingScheduling
{
    [Serializable]
    public class Lot
    {
        public string CustomerCode { get; set; }
        public string PartNumber { get; set; }
        public string AMO { get; set; }
        public string WorkOrderNumber { get; set; }
        public string FTStep { get; set; }
        public int TestTotalStep { get; set; }
        public string DeviceName { get; set; }
        public string Package { get; set; }
        public int LotQuantity { get; set; }
        public int Temperature { get; set; }
        public string Tester { get; set; }//assign tester
        public string Handler { get; set; }//assign handler
        public string LotNumber { get; set; }
        public DateTime TestReleaseDate { get; set; }
        public DateTime AssyInputDate { get; set; }
        public DateTime AssyPlanoutDate { get; set; }
        public DateTime TrackInDate { get; set; }
        public DateTime TrackInDate_Processing { get; set; }//0615 研究測試用
        public DateTime CompletionDate { get; set; }
        public DateTime DueDate { get; set; }
        public double CompletionTime { get; set; }
        public string CurrentStatus { get; set; }
        public double Weight { get; set; }
        public double ReleaseTime_of_Job { get; set; }
        public double DueTime_of_Job { get; set; }
    }

}
