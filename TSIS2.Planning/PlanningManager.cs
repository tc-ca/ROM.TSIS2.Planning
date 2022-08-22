using System;
using System.Configuration;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace TSIS2.Planning
{
    public class PlanningManager
    {
        public string Scope = "avsec";
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
            string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='ts_planningsettings'>
                                        <all-attributes />
	                                    <order attribute='ts_effectivedate' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='ts_taskstatus' operator='in'>
                                            <value>717750000</value>
                                          </condition>
                                          <condition attribute='statecode' value='0' operator='eq'/>
                                          <condition attribute='ts_effectivedate' value='" + DateTime.Now.ToString("yyyy-MM-dd") + @"' operator='on-or-before'/>
                                        </filter>
                                      </entity>
                                    </fetch>";

            EntityCollection planningSettings = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
            foreach (var planningSetting in planningSettings.Entities)
            {
                var task = planningSetting.GetAttributeValue<OptionSetValue>("ts_task").Value;
                var result = string.Empty;
                //Update task status to In Progress
                UpdatePlanningTaskStatus(svc, planningSetting, 717750001);

                switch (task)
                {
                    case 717750000: //Placeholder inspection
                        PlaceholderInspections placeholderInspections = new PlaceholderInspections();
                        result = placeholderInspections.GeneratePlaceHolderWorkOrders(svc, planningSetting);
                        break;
                }
                //Update task status to Completed
                UpdatePlanningTaskStatus(svc, planningSetting, 717750002);

                //Attach result as log file
                AttachLogFile(svc, result, planningSetting);
            }
        }

        private static void AttachLogFile(CrmServiceClient svc, string result, Entity planningSetting)
        {
            Entity Note = new Entity("annotation");
            Note["objectid"] = new EntityReference("ts_planningsettings", planningSetting.Id);
            Note["objecttypecode"] = "ts_planningsettings";
            Note["subject"] = "Function process log";
            Note["notetext"] = "Function process log file attached.";
            Note["filename"] = "Process.log";
            Note["mimetype"] = "text/plain";
            Note["documentbody"] = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(result));
            svc.Create(Note);
        }
        private static Entity UpdatePlanningTaskStatus(CrmServiceClient svc, Entity planningSetting, int taskStatus)
        {
            var planningSettingUpdate = new Entity("ts_planningsettings");
            planningSettingUpdate.Id = planningSetting.Id;
            planningSettingUpdate.Attributes["ts_taskstatus"] = new OptionSetValue(Convert.ToInt32(taskStatus));
            svc.Update(planningSettingUpdate);
            return planningSettingUpdate;
        }
        private void ISSOPlanning(CrmServiceClient svc)
        {
            TimeBasedPlanning timeBasedPlanning = new TimeBasedPlanning();
            //Security Plan Review (TDG)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["TDGSPR_IncidentTypeId"], 5);
            //Comprehensive Inspection (TDG)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["TDGCI_IncidentTypeId"], 5);
            //Security Plan Review (PAX)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["PAXSPR_IncidentTypeId"], 3);
            //Comprehensive Inspection (PAX)
            timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["PAXCI_IncidentTypeId"], 3);

            RiskBasedPlanning riskBasedPlanning = new RiskBasedPlanning();
            riskBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["TDGSI_IncidentTypeId"], ConfigurationManager.AppSettings["TDGVSI_IncidentTypeId"], "TDG");
            riskBasedPlanning.GenerateWorkOrderByIncidentType(svc, ConfigurationManager.AppSettings["PAXSI_IncidentTypeId"], ConfigurationManager.AppSettings["PAXOSI_IncidentTypeId"], "PAX");
        }
    }
}
