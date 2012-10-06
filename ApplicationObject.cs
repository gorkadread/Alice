using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Alice
{
    class ApplicationObject
    {
        public ApplicationObject()
        {
        }

        public ApplicationObject(string applicationName, string applicationPath)
        {
            ApplicationName = applicationName;
            ApplicationPath = applicationPath;
        }
        public string ApplicationName { get; set; }
        public string ApplicationPath { get; set; }
    }
}
