using System;
using System.Configuration;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;

namespace TSIS2.Planning
{
    public class PlanningManager
    {
        public static void Run()
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

                TimeBasedPlanning timeBasedPlanning = new TimeBasedPlanning();
                //Security Plan Review (TDG)
                timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["TDGSPR_IncidentTypeId"]);
            }
        }
    }
}
