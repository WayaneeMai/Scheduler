using System;
using System.Collections.Generic;
using System.Text;

namespace TestingScheduling
{
    [Serializable]
    public class ScheduleParameter
    {
        public List<MachineType> MachineTypes;
        
        public List<Resource> Testers_availablity;

        public List<Resource> Handlers_availablity;

        public List<AvailableAccessory> AvailableQuantity_accessory;

        public List<SecondardResource> SecondardResources;
    }
}
