using System;
using System.Linq;

namespace TSIS2.Planning
{
    internal class Program
    {
		static void Main(string[] args)
		{
			var scope = (args.Count() > 0) ? args[0] : "all";

			PlanningManager planManager = new PlanningManager(scope);
			planManager.Run();
			Console.WriteLine("Press any key to exit.");
			Console.ReadLine();
		}
    }
}
