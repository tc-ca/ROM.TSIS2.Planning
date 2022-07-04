using System;
using System.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using NLog;

namespace TSIS2.Planning
{
    public class TimeBasedPlanning
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Generate work orders by incident type
        /// </summary>
        /// <param name="svc"></param>
        /// <param name="incidentTypeId"></param>
        /// <param name="lastYears"></param>
        public void GenerateWorkOrderByIncidentType(CrmServiceClient svc, string incidentTypeId, int lastYears)
        {
            try
            {
                logger.Info("Start processing by incident type id {0}", incidentTypeId);
                string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name'/>
                        <attribute name='ts_site'/>
                        <attribute name='ts_region'/>
                        <attribute name='msdyn_primaryincidenttype'/>
                        <attribute name='ovs_operationtypeid'/>
                        <attribute name='ovs_operationid' />
                        <attribute name='msdyn_serviceaccount'/>
                        <attribute name='msdyn_workorderid'/>
                        <attribute name='msdyn_functionallocation'/>
                        <order attribute='msdyn_name' descending='false'/>
                        <filter type='and'>
                        <condition attribute='msdyn_primaryincidenttype' operator='eq' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/>
                        <condition attribute='statecode' value='0' operator='eq'/>
                        </filter>
                        <link-entity name='tc_tcfiscalyear' from='tc_tcfiscalyearid' to='ovs_fiscalyear' link-type='inner' alias='ag'>
                        <filter type='and'>
                        <condition attribute='tc_fiscalstart' operator='last-x-years' value='" + lastYears + @"'/>
                        </filter>
                        </link-entity>
                        </entity>
                        </fetch>";

                EntityCollection workorders = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
                logger.Info("In total there are {0} work order(s) for incident type id {1}.", workorders.Entities.Count, incidentTypeId);
                foreach (var c in workorders.Entities)
                {
                    logger.Info("Work Order Name: {0}, Id {1}", c.Attributes["msdyn_name"], c.Id);
                }

                fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='ovs_operationtype'>
                        <attribute name='ovs_operationtypeid'/>
                        <attribute name='ovs_name'/>
                        <order attribute='ovs_name' descending='false'/>
                        <link-entity name='ts_ovs_operationtypes_msdyn_incidenttypes' from='ovs_operationtypeid' to='ovs_operationtypeid' visible='false' intersect='true'>
                        <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' alias='ab'>
                        <filter type='and'>
                        <condition attribute='msdyn_incidenttypeid' operator='eq'  uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/>
                        <condition attribute='statecode' value='0' operator='eq'/>
                        </filter>
                        </link-entity>
                        </link-entity>
                        </entity>
                        </fetch>";

                EntityCollection operationtypes = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
                logger.Info("In total there are {0} Operation Types that related to incident type id {1}", operationtypes.Entities.Count, incidentTypeId);
                foreach (var operationType in operationtypes.Entities)
                {
                    logger.Info("Operation type Name: {0}", operationType.Attributes["ovs_name"]);
                }

                //All operations related to incident type 
                fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='ovs_operation'>
                    <attribute name='ovs_name' />
                    <attribute name='createdon' />
                    <attribute name='ovs_operationtypeid' />
                    <attribute name='ovs_operationid' />
                    <attribute name='ts_stakeholder' />
                    <attribute name='ts_site' />
                    <order attribute='ovs_name' descending='false' />
                    <link-entity name='ovs_operationtype' from='ovs_operationtypeid' to='ovs_operationtypeid' link-type='inner' alias='ac'>
                      <link-entity name='ts_ovs_operationtypes_msdyn_incidenttypes' from='ovs_operationtypeid' to='ovs_operationtypeid' visible='false' intersect='true'>
                        <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' alias='ad'>
                          <filter type='and'>
                            <condition attribute='msdyn_incidenttypeid' operator='eq' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"' />
                          </filter>
                        </link-entity>
                      </link-entity>
                    </link-entity>
                    <link-entity name='msdyn_functionallocation' from='msdyn_functionallocationid' to='ts_site' visible='false' link-type='outer' alias='siteregion'>
                      <attribute name='ts_region' />
                    </link-entity>
                    <filter type='and'>
                        <condition attribute='statecode' value='0' operator='eq'/>
                    </filter>
                  </entity>
                </fetch>";

                EntityCollection operations = svc.RetrieveMultiple(new FetchExpression(fetchQuery));
                logger.Info("In total there are {0} Operations that related to incident type id {1}.", operations.Entities.Count, incidentTypeId);
                foreach (var operation in operations.Entities)
                {
                    bool woExists = false;
                    foreach (var workOrderCandidate in workorders.Entities)
                    {
                        if (workOrderCandidate.GetAttributeValue<EntityReference>("msdyn_serviceaccount").Id == operation.GetAttributeValue<EntityReference>("ts_stakeholder").Id &&
                            (workOrderCandidate.GetAttributeValue<EntityReference>("ts_site") != null && workOrderCandidate.GetAttributeValue<EntityReference>("ts_site").Id == operation.GetAttributeValue<EntityReference>("ts_site").Id) &&
                            (workOrderCandidate.GetAttributeValue<EntityReference>("ts_region") != null && workOrderCandidate.GetAttributeValue<EntityReference>("ts_region").Id == ((EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value).Id) &&
                            (workOrderCandidate.GetAttributeValue<EntityReference>("msdyn_primaryincidenttype") != null && workOrderCandidate.GetAttributeValue<EntityReference>("msdyn_primaryincidenttype").Id == new Guid(incidentTypeId)) &&
                            (workOrderCandidate.GetAttributeValue<EntityReference>("ovs_operationtypeid") != null && workOrderCandidate.GetAttributeValue<EntityReference>("ovs_operationtypeid").Id == operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Id) &&
                            (workOrderCandidate.GetAttributeValue<EntityReference>("ovs_operationid") != null && workOrderCandidate.GetAttributeValue<EntityReference>("ovs_operationid").Id == operation.Id))
                        {
                            logger.Info("Work order {0}, id {1} exists for operation: {2}, stake holder: {3}, type: {4}, site {5}, region {6} ",
                                workOrderCandidate.Attributes["msdyn_name"],
                                workOrderCandidate.Id,
                                operation.Attributes["ovs_name"],
                                operation.GetAttributeValue<EntityReference>("ts_stakeholder").Name,
                                operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Name,
                                operation.GetAttributeValue<EntityReference>("ts_site").Name,
                                ((EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value).Name);
                            woExists = true;
                            break;
                        }
                    }
                    if (!woExists)
                    {
                        logger.Info("Create planning work order for operation: {0}, stake holder: {1}, type: {2}, site {3}, region {4} ",
                        operation.Attributes["ovs_name"],
                        operation.GetAttributeValue<EntityReference>("ts_stakeholder").Name,
                        operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Name,
                        operation.GetAttributeValue<EntityReference>("ts_site").Name,
                        ((EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value).Name);

                        Entity workOrder = new Entity("msdyn_workorder");
                        workOrder["msdyn_serviceaccount"] = operation.GetAttributeValue<EntityReference>("ts_stakeholder");
                        workOrder["ovs_operationid"] = operation.ToEntityReference();
                        workOrder["ovs_operationtypeid"] = operation.GetAttributeValue<EntityReference>("ovs_operationtypeid");
                        workOrder["ts_site"] = operation.GetAttributeValue<EntityReference>("ts_site");
                        workOrder["ts_region"] = (EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value;
                        workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(incidentTypeId));
                        var tradeName = Utilities.GetDefaultTradeNameByStakeHolder(svc, operation.GetAttributeValue<EntityReference>("ts_stakeholder").Id.ToString());
                        if (tradeName != null)
                        {
                            workOrder["ts_tradenameid"] = tradeName;
                        }
                        workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(ConfigurationManager.AppSettings["RationalePlannedId"]));  //Planned
                        Guid workOrderId = svc.Create(workOrder);
                        logger.Info("New Work Order Id: {0}", workOrderId);
                    }
                }

                //Enable the follow line if want to delete generated test work orders.
                //Utilities.DeleteWorkOrders(svc, incidentTypeId, workorders);
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
    }
}
