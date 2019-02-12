using System;

namespace RevitCommon.Attributes
{
    /// <summary>
    /// This attribute is for HKS tools only to be used in conjunction with the CommandManager, a tool that allows users to add/remove HKS standard tools if they so desire.
    /// This Attribute is specifically for IExternalApplication plugins.
    /// </summary>
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

    /// <summary>
    /// This attribute is for HKS tools only to be used in conjunction with the CommandManager, a tool that allows users to add/remove HKS standard tools if they so desire.
    /// This Attribute is specifically for IExternalCommand plugins. Probably not going to be used frequently.
    /// </summary>
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