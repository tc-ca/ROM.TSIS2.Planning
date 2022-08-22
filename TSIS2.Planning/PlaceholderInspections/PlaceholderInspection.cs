using System;
using System.Configuration;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;

namespace TSIS2.Planning
{
    public class PlaceholderInspections
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public string GeneratePlaceHolderWorkOrders(CrmServiceClient svc, Entity planningSetting)
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
                    workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(ConfigurationManager.AppSettings["RationaleInitDraftId"]));  //Draft
                    Guid workOrderId = svc.Create(workOrder);
                    sb.AppendLine(String.Format("Created New Work Order Id {0}", workOrderId));
                }
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
