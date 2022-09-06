using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Azure.WebJobs.Host;
using System.Text;

namespace TSIS2.PlanningFunction
{
    public class TimeBasedPlanning
    {
        /// <summary>
        /// Generate work orders by incident type
        /// </summary>
        /// <param name="svc"></param>
        /// <param name="incidentTypeId"></param>
        /// <param name="lastYears"></param>
        /// <param name="logger"></param>
        /// <param name="activityType"></param>
        /// <param name="HQId"></param>
        /// <returns></returns>
        public string GenerateWorkOrderByIncidentType(CrmServiceClient svc, string incidentTypeId, int lastYears, TraceWriter logger, ActivityType activityType, string HQId="")
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                logger.Info(String.Format("Start processing by incident type id {0}", incidentTypeId));
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
                sb.AppendLine(String.Format("Find {0} work order(s) for incident type id {1} for past {2} years.", workorders.Entities.Count, incidentTypeId, lastYears));
                foreach (var c in workorders.Entities)
                {
                    sb.AppendLine(String.Format("Work Order Name: {0}, Id {1}", c.Attributes["msdyn_name"], c.Id));
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
                sb.AppendLine(String.Format("Find {0} Operation Type(s) that related to incident type id {1}", operationtypes.Entities.Count, incidentTypeId)); ;
                foreach (var operationType in operationtypes.Entities)
                {
                    sb.AppendLine(String.Format("Operation type Name: {0}", operationType.Attributes["ovs_name"]));
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
                    <attribute name='ts_typeofdangerousgoods' />
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
                sb.AppendLine(String.Format("Find {0} Operations that related to incident type id {1}.", operations.Entities.Count, incidentTypeId));
                foreach (var operation in operations.Entities)
                {
                    bool canBePlanned = false;
                    bool nonSchedule1Apply = false;
                    if (string.IsNullOrWhiteSpace(HQId))
                    {
                        canBePlanned = true;
                    }
                    else
                    {
                        canBePlanned = HQId.IndexOf(operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Id.ToString(), StringComparison.InvariantCultureIgnoreCase) >= 0;
                    }
                    if (activityType == ActivityType.TDGSPR  || activityType == ActivityType.TDGCI)
                    {
                        if (Utilities.CheckIfHQExists(svc, operation.GetAttributeValue<EntityReference>("ts_stakeholder").Id.ToString(), HQId))
                        {
                            canBePlanned = true;
                        }
                        else
                        {
                            var dangerousgoodsType = operation.GetAttributeValue<OptionSetValue>("ts_typeofdangerousgoods")?.Value;
                            //If undecided or schedule 1 dangerous goods
                            if (dangerousgoodsType == null || dangerousgoodsType == 717750001)
                            {
                                canBePlanned = true;
                            }
                            if (activityType == ActivityType.TDGCI)
                            {
                                //Non-Schedule 1 Dangerous Goods
                                if (dangerousgoodsType == 717750002)
                                {
                                    nonSchedule1Apply = true;
                                }
                            }
                        }
                    }
                    if (canBePlanned || nonSchedule1Apply)
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
                                sb.AppendLine(String.Format("Work order {0}, id {1} exists for operation: {2}, stake holder: {3}, type: {4}, site {5}, region {6} ",
                                    workOrderCandidate.Attributes["msdyn_name"],
                                    workOrderCandidate.Id,
                                    operation.Attributes["ovs_name"],
                                    operation.GetAttributeValue<EntityReference>("ts_stakeholder").Name,
                                    operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Name,
                                    operation.GetAttributeValue<EntityReference>("ts_site").Name,
                                    ((EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value).Name));
                                woExists = true;
                                break;
                            }
                        }
                        if (!woExists)
                        {
                            sb.AppendLine(String.Format("Create planning work order for operation: {0}, stake holder: {1}, type: {2}, site {3}, region {4} ",
                            operation.Attributes["ovs_name"],
                            operation.GetAttributeValue<EntityReference>("ts_stakeholder").Name,
                            operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Name,
                            operation.GetAttributeValue<EntityReference>("ts_site").Name,
                            ((EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value).Name));

                            Entity workOrder = new Entity("msdyn_workorder");
                            workOrder["msdyn_serviceaccount"] = operation.GetAttributeValue<EntityReference>("ts_stakeholder");
                            workOrder["ovs_operationid"] = operation.ToEntityReference();
                            workOrder["ovs_operationtypeid"] = operation.GetAttributeValue<EntityReference>("ovs_operationtypeid");
                            workOrder["ts_site"] = operation.GetAttributeValue<EntityReference>("ts_site");
                            workOrder["ts_region"] = (EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value;
                            if (nonSchedule1Apply)
                            {
                                workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(Environment.GetEnvironmentVariable("ROM_TDGNS1CI_IncidentTypeId", EnvironmentVariableTarget.Process)));
                            }
                            else
                            {
                                workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(incidentTypeId));
                            }
                            var tradeName = Utilities.GetDefaultTradeNameByStakeHolder(svc, operation.GetAttributeValue<EntityReference>("ts_stakeholder").Id.ToString());
                            if (tradeName != null)
                            {
                                workOrder["ts_tradenameid"] = tradeName;
                            }
                            workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(Environment.GetEnvironmentVariable("ROM_Category_PlannedId", EnvironmentVariableTarget.Process)));  //Planned
                            workOrder["ts_state"] = new OptionSetValue(Convert.ToInt32(717750000));   //Draft
                            workOrder["ts_origin"] = String.Format("Forecast {0}/{1}", (DateTime.Now.AddYears(1)).ToString("yyyy"), (DateTime.Now.AddYears(2)).ToString("yy"));
                            Guid workOrderId = svc.Create(workOrder);
                            sb.AppendLine(String.Format("New Work Order Id: {0}", workOrderId));
                        }
                    }
                }
                sb.AppendLine(String.Format("Processing by Incident Type {0} Completed Successfully.", incidentTypeId));
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
