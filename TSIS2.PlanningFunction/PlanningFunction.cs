using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using TSIS2.PlanningFunction.Planning;

namespace TSIS2.PlanningFunction
{
    public static class PlanningFunction
    {
        [FunctionName("TSIS2PlanningFunction")]
        public static void Run([TimerTrigger("0 */5 * * * *"
            #if DEBUG
            , RunOnStartup=true
            #endif
            )]TimerInfo myTimer, TraceWriter log)
        {
            log.Info($"TSIS2PlanningFunction Timer trigger function executed at: {DateTime.Now}");
            Process(log);
        }

        private static void Process(TraceWriter log)
        {
            string url = Environment.GetEnvironmentVariable("ROM_Url", EnvironmentVariableTarget.Process); 
            string clientId = Environment.GetEnvironmentVariable("ROM_ClientId", EnvironmentVariableTarget.Process); 
            string clientSecret = Environment.GetEnvironmentVariable("ROM_ClientSecret", EnvironmentVariableTarget.Process);
            string connectString = $"AuthType=ClientSecret;url={url};ClientId={clientId};ClientSecret={clientSecret}";
            try
            {
                using (var svc = new CrmServiceClient(connectString))
                {
                    //This is test code.
                    //WhoAmIRequest request = new WhoAmIRequest();
                    //WhoAmIResponse response = (WhoAmIResponse)svc.Execute(request);
                    //log.Info("log UserId is " + response.UserId.ToString());

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
                    log.Info(String.Format("In total there are {0} planning settings", planningSettings.Entities.Count));
                    foreach (var planningSetting in planningSettings.Entities)
                    {
                        var task = planningSetting.GetAttributeValue<OptionSetValue>("ts_task").Value;
                        var result = String.Format("Start Processing Planning Task {0} - {1} at {2} " + Environment.NewLine, planningSetting.Id, planningSetting.FormattedValues["ts_task"], DateTime.Now.ToUniversalTime()
                         .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'"));
                        //Update task status to In Progress
                        UpdatePlanningTaskStatus(svc, planningSetting, 717750001);
                        TimeBasedPlanning timeBasedPlanning = new TimeBasedPlanning();
                        switch (task)
                        {
                            case 717750000: //Placeholder inspection
                                PlaceholderInspections placeholderInspections = new PlaceholderInspections();
                                result += placeholderInspections.GeneratePlaceHolderWorkOrders(svc, planningSetting, log);
                                break;
                            case 717750001: //PAX SP/SRA Review
                                result += timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, Environment.GetEnvironmentVariable("ROM_PAXSPR_IncidentTypeId", EnvironmentVariableTarget.Process), 3, log, ActivityType.PAXSPR);
                                break;
                            case 717750002: //PAX Comprehensive Inspection - Large PAX
                                result += timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, Environment.GetEnvironmentVariable("ROM_PAXCI_IncidentTypeId", EnvironmentVariableTarget.Process), 3, log, ActivityType.PAXCI, Environment.GetEnvironmentVariable("ROM_PAX_HQ_Large_Id", EnvironmentVariableTarget.Process));
                                break;
                            case 717750003: //PAX Comprehensive Inspection - Small PAX
                                result += timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, Environment.GetEnvironmentVariable("ROM_PAXCI_IncidentTypeId", EnvironmentVariableTarget.Process), 5, log, ActivityType.PAXCI, Environment.GetEnvironmentVariable("ROM_PAX_HQ_Small_Id", EnvironmentVariableTarget.Process));
                                break;
                            case 717750005: //TDG Security Plan Review
                                result += timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, Environment.GetEnvironmentVariable("ROM_TDGSPR_IncidentTypeId", EnvironmentVariableTarget.Process), 5, log, ActivityType.TDGSPR, Environment.GetEnvironmentVariable("ROM_TDG_HQ_Id", EnvironmentVariableTarget.Process));
                                break;
                            case 717750006: //TDG Comprehensive Inspection
                                result += timeBasedPlanning.GenerateWorkOrderByIncidentType(svc, Environment.GetEnvironmentVariable("ROM_TDGCI_IncidentTypeId", EnvironmentVariableTarget.Process), 5, log, ActivityType.TDGCI, Environment.GetEnvironmentVariable("ROM_TDG_HQ_Id", EnvironmentVariableTarget.Process));
                                break;
                            case 717750004: //TDG Site Inspection
                                RiskBasedPlanning riskBasedPlanning = new RiskBasedPlanning();
                                result += riskBasedPlanning.GenerateWorkOrderByIncidentType(svc, Environment.GetEnvironmentVariable("ROM_TDGSI_IncidentTypeId"), Environment.GetEnvironmentVariable("ROM_TDGVSI_IncidentTypeId"), "TDG",log);
                                break;
                        }
                        //Update task status to Completed
                        UpdatePlanningTaskStatus(svc, planningSetting, 717750002);

                        //Attach result as log file
                        AttachLogFile(svc, result, planningSetting);
                    }
                }
            }
            catch (Exception e)
            {
                log.Error(e.Message);
            }
        }

        private static void AttachLogFile(CrmServiceClient svc, string result, Entity planningSetting)
        {
            Entity Note = new Entity("annotation");
            Note["objectid"] = new EntityReference("ts_planningsettings", planningSetting.Id);
            Note["objecttypecode"] = "ts_planningsettings";
            Note["subject"] = "Planning Function process log";
            Note["notetext"] = "Log file attached.";
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
    }
}
