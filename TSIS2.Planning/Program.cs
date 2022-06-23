using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace TSIS2.Planning
{
    internal class Program
    {
        static void Main(string[] args)
        {
            PlanningManager.Run();
            Console.WriteLine("Press any key to exit.");
            Console.ReadLine();
        }
    }
}
