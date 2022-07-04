using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;
using System.Linq;

namespace TSIS2.Planning
{
    internal static class Utilities
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Clean up generated test work orders
        /// </summary>
        /// <param name="svc"></param>
        /// <param name="incidentTypeId"></param>
        /// <param name="workordersToKeep"></param>
        public static void DeleteWorkOrders(CrmServiceClient svc, string incidentTypeId, EntityCollection workordersToKeep)
        {
            string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name'/>
                        <attribute name='ts_site'/>
                        <attribute name='ts_region'/>
                        <attribute name='msdyn_primaryincidenttype'/>
                        <attribute name='ovs_operationtypeid'/>
                        <attribute name='msdyn_serviceaccount'/>
                        <attribute name='msdyn_workorderid'/>
                        <attribute name='msdyn_functionallocation'/>
                        <order attribute='msdyn_name' descending='false'/>
                        <filter type='and'>
                        <condition attribute='msdyn_primaryincidenttype' operator='eq' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/>
                        </filter>
                        </entity>
                        </fetch>";
            EntityCollection workordersToBeDelete = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
            logger.Info("In total there are {0} work order(s) for incident type id {1}.", workordersToBeDelete.Entities.Count, incidentTypeId);
            foreach (var workorderToBeDelete in workordersToBeDelete.Entities)
            {
                if (workordersToKeep == null || (workordersToKeep.Entities.Count > 0 && !workordersToKeep.Entities.Any(a => a.Id == workorderToBeDelete.Id)))
                {
                    logger.Info("Delete Work Order Name: {0}, Id {1}", workorderToBeDelete.Attributes["msdyn_name"], workorderToBeDelete.Id);
                    svc.Delete("msdyn_workorder", workorderToBeDelete.Id);
                }
            }
        }

        /// <summary>
        /// Get first trade name by stake holder Id.
        /// </summary>
        /// <param name="svc"></param>
        /// <param name="stakeholderId"></param>
        /// <returns></returns>
        public static EntityReference GetDefaultTradeNameByStakeHolder(CrmServiceClient svc, string stakeholderId)
        {
            string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='ts_tradename'>
                                <attribute name='ts_tradenameid' />
                                <attribute name='ts_name' />
                                <attribute name='createdon' />
                                <order attribute='ts_name' descending='false' />
                                <filter type='and'>
                                  <condition attribute='ts_stakeholderid' operator='eq' uitype='account' value='" + stakeholderId + @"' />
                                  <condition attribute='statecode' value='0' operator='eq'/>
                                </filter>
                              </entity>
                            </fetch>";
            EntityCollection tradenames = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
            //logger.Info("In total there are {0} tradename for stakeholder id {1}.", tradenames.Entities.Count, stakeholderId);
            //foreach (var tradename in tradenames.Entities)
            //{
            //    logger.Info("tradename Name: {0}", tradename.Attributes["ts_name"]);
            //}
            if (tradenames.Entities != null && tradenames.Entities.Count > 0)
            {
                return tradenames.Entities[0].ToEntityReference();
            }
            else
            {
                return null;
            }
        }
    }
}
