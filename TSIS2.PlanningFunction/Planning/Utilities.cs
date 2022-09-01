using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace TSIS2.PlanningFunction
{
    public static class Utilities
    {
        /// <summary>
        /// Get Default TradeName By StakeHolder
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
            if (tradenames.Entities != null && tradenames.Entities.Count > 0)
            {
                return tradenames.Entities[0].ToEntityReference();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Check if HQ operation type exists for stake holder's operations
        /// </summary>
        /// <param name="svc"></param>
        /// <param name="stakeholderId"></param>
        /// <param name="HQIds"></param>
        /// <returns></returns>
        public static bool CheckIfHQExists(CrmServiceClient svc, string stakeholderId, string HQIds)
        {
            string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='ovs_operation'>
                        <attribute name='ovs_name' />
                        <attribute name='ovs_operationtypeid' />
                        <attribute name='ovs_operationid' />
                        <filter type='and'>
                          <condition attribute='ts_stakeholder' operator='eq' value='" + stakeholderId + @"' />
                        </filter>
                      </entity>
                    </fetch>";
            bool HQExists = false;
            EntityCollection operations = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
            if (operations.Entities != null && operations.Entities.Count > 0)
            {
                var HQOperations = operations.Entities.Where(op => HQIds.ToLower().Contains(op.GetAttributeValue<EntityReference>("ovs_operationtypeid").Id.ToString().ToLower())).ToList();
                if (HQOperations.Count() > 0)
                {
                    HQExists = true;
                }
            }
            return HQExists;
        }
    }
}
