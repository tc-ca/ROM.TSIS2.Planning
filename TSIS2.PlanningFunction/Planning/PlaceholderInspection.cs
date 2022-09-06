using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Azure.WebJobs.Host;
using System.Text;

namespace TSIS2.PlanningFunction
{
    public class PlaceholderInspections
    {
        public string GeneratePlaceHolderWorkOrders(CrmServiceClient svc, Entity planningSetting, TraceWriter logger)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                var totalCount = planningSetting.GetAttributeValue<int>("ts_totalcount");
                for (var i = 0; i < totalCount; i++)
                {
                    Entity workOrder = new Entity("msdyn_workorder");
                    workOrder["msdyn_serviceaccount"] = planningSetting.GetAttributeValue<EntityReference>("ts_stakeholder");
                    if (planningSetting.GetAttributeValue<EntityReference>("ts_region") != null)
                    {
                        workOrder["ts_region"] = planningSetting.GetAttributeValue<EntityReference>("ts_region");
                    }
                    if (planningSetting.GetAttributeValue<EntityReference>("ts_workordertype") != null)
                    {
                        workOrder["msdyn_workordertype"] = planningSetting.GetAttributeValue<EntityReference>("ts_workordertype");
                    }
                    if (planningSetting.GetAttributeValue<EntityReference>("ts_workorderowner") != null)
                    {
                        workOrder["ownerid"] = planningSetting.GetAttributeValue<EntityReference>("ts_workorderowner");
                    }

                    workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(Environment.GetEnvironmentVariable("ROM_Category_PlannedId", EnvironmentVariableTarget.Process)));  //Planned
                    workOrder["ts_state"] = new OptionSetValue(Convert.ToInt32(717750000));   //Draft
                    workOrder["ts_origin"] = String.Format("Forecast {0}/{1}", (DateTime.Now.AddYears(1)).ToString("yyyy"), (DateTime.Now.AddYears(2)).ToString("yy"));

                    Guid workOrderId = svc.Create(workOrder);
                    sb.AppendLine(String.Format("Created New Work Order Id {0}", workOrderId));
                }
                sb.AppendLine("Processing Batch Create Work Orders Completed Successfully.");
            }
            catch (Exception ex)
            {
                sb.AppendLine(ex.Message);
                logger.Error(ex.Message);
            }
            return sb.ToString();
        }
    }
}
