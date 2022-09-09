using System;
using System.Collections.Generic;
using System.Text;

namespace TestingScheduling
{
    [Serializable]
    public class Job_Operation_Index:Lot
    {
        public int JobIndex { get; set; }//jobIndex
        public int OperationIndex { get; set; }//operation index. i.e.,step
        public int MachineTypeIndex { get; set; }
        public int FamilyIndex { get; set; }//device index
        public double Time { get; set; }//to save start time and finish time
    }
}
