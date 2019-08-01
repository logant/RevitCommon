using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitCommon
{
    [Obsolete("Temporarily replaced HKSTools.WriteToHome function to provide backwards compatibility for a transition period")]
    public class HKS
    {
        public static void WriteToHome(string commandName, string appVersion, string userName)
        {
            // Do nothing
        }
    }
}
