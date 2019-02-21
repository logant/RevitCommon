/* 
* Copyright 2012 © Victor Chekalin
* 
* THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
* KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
* IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
* PARTICULAR PURPOSE.
* https://github.com/chekalin-v/VCExtensibleStorageExtension
* http://thebuildingcoder.typepad.com/blog/2013/05/effortless-extensible-storage.html
*/

using System;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace RevitCommon.Attributes
{

    [AttributeUsage(AttributeTargets.Class)]
    public class SchemaAttribute : Attribute
    {
        private readonly string _schemaName;
        private readonly Guid _guid;

        public SchemaAttribute(string guid, string schemaName)
        {
            _schemaName = schemaName;
            _guid = new Guid(guid);
        }

        public string SchemaName
        {
            get { return _schemaName; }
        }

        public Guid ApplicationGUID { get; set; }

        public string Documentation { get; set; }

        public Guid GUID
        {
            get { return _guid; }
        }

        public AccessLevel ReadAccessLevel { get; set; }

        public AccessLevel WriteAccessLevel { get; set; }

        public string VendorId { get; set; }
    }
}