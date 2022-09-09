using System;
using System.Collections.Generic;
using System.Text;

namespace TestingScheduling
{
    [Serializable]
    public class Resource: Job_Operation_Index
    {
        public string ResourceName{ get; set; }
        public int ResourceIndex { get; set; }
        public double AvailableTime { get; set; }
        public string ResourceType { get; set; }
        public int RequireQuantity { get; set; }
        public string ResourceStatus { get; set; }
        public string ResourceSubStatus { get; set; }
        public int ResourceQuantity { get; set; }
        public string ResourceLocation { get; set; }
        public DateTime EstimationAvailableTime { get; set; }
    }

    [Serializable]
    public class MachineType: Job_Operation_Index
    {
        public int TesterIndex { get; set; }
        public int HandlerIndex { get; set; }
        public double AvailableTime { get; set; }
    }

    [Serializable]
    public class AvailableAccessory
    {
        public string AccessoryType { get; set; }
        public string AccessoryName { get; set; }
        public int Accessory_Type_Index { get; set; }
        public int AvailableAccessoryIndex { get; set; }
        public int AvailableQuantity { get; set; }
        public double Time { get; set; }
        public double WaitingTime { get; set; }
    }

    public class SecondardResource
    {
        public int PrimaryResourceIndex { get; set; }
        public string PrimaryResourceType { get; set; }
        public string PrimaryResourceName { get; set; }
        public int SecondardResourceIndex { get; set; }
        public string SecondaryResourceType { get; set; }
        public string SecondaryResourceName { get; set; }
    }

}
