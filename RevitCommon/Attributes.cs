using System;

namespace RevitCommon.Attributes
{
    public class ExtApp : Attribute
    {
        public string Name { get; set; }

        public string Guid { get; set; }

        public string Description { get; set; }

        public string Vendor { get; set; }

        public string VendorDescription { get; set; }

        public string[] Commands { get; set; }

        public bool ForceEnabled { get; set; }
    }

    public class ExtCmd : Attribute
    {
        public string Name { get; set; }

        public string Guid { get; set; }

        public string Description { get; set; }

        public string Vendor { get; set; }

        public string VendorDescription { get; set; }

        public string[] Commands { get; set; }

        public bool ForceEnabled { get; set; }
    }
}