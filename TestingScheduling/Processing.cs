using System;
using System.Collections.Generic;
using System.Text;

namespace TestingScheduling
{
    public class Processing:MachineType
    {
        public string TesterType { get; set; }
        public string HandlerType { get; set; }
        public double UnitProcessingPerHour { get; set; }
        public double ProcessingTime { get; set; }
        public double SetupTime { get; set; }
        public string IsDefault { get; set; }
    }
}
