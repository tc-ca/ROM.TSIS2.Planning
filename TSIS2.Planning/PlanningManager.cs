using System.Configuration;
using Microsoft.Xrm.Tooling.Connector;

namespace TSIS2.Planning
{
    public class PlanningManager
    {
        public string Scope = "all";
        public PlanningManager(string scope)
        {
            Scope = scope;
        }
        public void Run()
        {
            string url = ConfigurationManager.AppSettings["Url"];
            string clientId = ConfigurationManager.AppSettings["ClientId"];
            string clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
            string connectString = $"AuthType=ClientSecret;url={url};ClientId={clientId};ClientSecret={clientSecret}";
            using (var svc = new CrmServiceClient(connectString))
            {
                //This is test code.
                //WhoAmIRequest request = new WhoAmIRequest();
                //WhoAmIResponse response = (WhoAmIResponse)svc.Execute(request);
                //Console.WriteLine("UserId is {0}", response.UserId);

                switch (Scope.ToLower())
                {
                    case "avsec":
                        AvSecPlanning(svc);
                        break;

                    case "isso":
                        ISSOPlanning(svc);
                        break;

                    default:
                        AvSecPlanning(svc);
                        ISSOPlanning(svc);
                        break;
                }              
            }
        }

        private void AvSecPlanning(CrmServiceClient svc)
        {

        }

        private void ISSOPlanning(CrmServiceClient svc)
        {
            TimeBasedPlanning timeBasedPlanning = new TimeBasedPlanning();
            //Security Plan Review (TDG)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["TDGSPR_IncidentTypeId"], 5);
            //Comprehensive Inspection (TDG)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["TDGCI_IncidentTypeId"], 5);
            //Security Plan Review (PAX)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["PAXSPR_IncidentTypeId"],3);
            //Comprehensive Inspection (PAX)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["PAXCI_IncidentTypeId"],3);
        }
    }
}
