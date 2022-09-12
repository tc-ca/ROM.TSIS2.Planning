using Microsoft.Azure.WebJobs.Host;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSIS2.PlanningFunction.Planning
{
    public class RiskBasedPlanning
    {
        public string GenerateWorkOrderByIncidentType(CrmServiceClient svc, string incidentTypeId, string altIncidentTypeId, string altFlag, TraceWriter logger, string altSecondIncidentTypeId = "")
        {
            StringBuilder stringBuilder = new StringBuilder();
            string logMessage = "";
            try
            {
                logMessage = String.Format("Start processing by incident type id {0}", incidentTypeId);
                logger.Info(logMessage);
                stringBuilder.AppendLine(logMessage);

                string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                  <entity name='ts_riskcategory'>
                    <attribute name='ts_name' />
                    <attribute name='ts_riskcategoryid' />
                    <attribute name='ts_riskscoreminimum' />
                    <attribute name='ts_riskscoremaximum' />
                    <attribute name='ts_operationtype' />
                    <attribute name='ts_interval' />
                    <attribute name='ts_frequency' />
                    <order attribute='ts_name' descending='false' />
                    <filter type='and'>
                      <condition attribute='statuscode' operator='eq' value='1' />
                    </filter>
                    <link-entity name='ovs_operationtype' from='ovs_operationtypeid' to='ts_operationtype' link-type='inner' alias='ai'>
                      <link-entity name='ts_ovs_operationtypes_msdyn_incidenttypes' from='ovs_operationtypeid' to='ovs_operationtypeid' visible='false' intersect='true'>
                        <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' alias='aj'>
                          <filter type='and'>
                            <condition attribute='msdyn_incidenttypeid' operator='eq' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/> 
                            <condition attribute='statecode' value='0' operator='eq'/>
                          </filter>
                        </link-entity>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

                EntityCollection riskThresholds = svc.RetrieveMultiple(new FetchExpression(fetchQuery));

                logMessage = String.Format("In total there are {0} Risk Thresholds that related to incident type id {1}", riskThresholds.Entities.Count, incidentTypeId);
                logger.Info(logMessage);
                stringBuilder.AppendLine(logMessage);

                foreach (var riskThreshold in riskThresholds.Entities)
                {
                    logMessage = String.Format("Risk Thresholds: {0}, {1}, {2}, {3}",
                        riskThreshold.GetAttributeValue<EntityReference>("ts_operationtype").Name,
                        riskThreshold.Attributes["ts_name"],
                        riskThreshold.GetAttributeValue<int>("ts_riskscoreminimum"),
                        riskThreshold.GetAttributeValue<int>("ts_riskscoremaximum"));

                    logger.Info(logMessage);
                    stringBuilder.AppendLine(logMessage);
                }

                fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='ovs_operation'>
                    <attribute name='ovs_name' />
                    <attribute name='createdon' />
                    <attribute name='ovs_operationtypeid' />
                    <attribute name='ovs_operationid' />
                    <attribute name='ts_stakeholder' />
                    <attribute name='ts_site' />
                    <attribute name='ts_riskscore' />
                    <attribute name='ts_visualsecurityinspection' />
                    <attribute name='ts_issecurityinspectionsite' />
                    <attribute name='ts_typeofdangerousgoods' />
                    <attribute name='ts_stationtype' />
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

                logMessage = String.Format("In total there are {0} Operations that related to incident type id {1}.", operations.Entities.Count, incidentTypeId);
                logger.Info(logMessage);
                stringBuilder.AppendLine(logMessage);

                foreach (var operation in operations.Entities)
                {
                    int riskScore = operation.GetAttributeValue<int>("ts_riskscore");
                    bool vsi = false;
                    bool si = false;
                    //717,750,001 - Yes
                    if (altFlag.Equals("TDG", StringComparison.InvariantCultureIgnoreCase))
                    {
                        if (operation.GetAttributeValue<OptionSetValue>("ts_visualsecurityinspection") != null &&
                            operation.GetAttributeValue<OptionSetValue>("ts_visualsecurityinspection").Value == 717750001)
                        {
                            vsi = true;
                        }
                    }
                    else if(altFlag.Equals("PAX", StringComparison.InvariantCultureIgnoreCase))
                    {
                        //if this is a site where Security Inspections (SIs) are conducted
                        if (operation.GetAttributeValue<OptionSetValue>("ts_issecurityinspectionsite") != null && operation.GetAttributeValue<OptionSetValue>("ts_issecurityinspectionsite").Value == 717750001)
                        {
                            vsi = false;
                            si = true;
                        }
                        else
                        {
                            vsi = false;
                            si = false;
                        }
                    }
                    else
                    {
                        if (operation.GetAttributeValue<OptionSetValue>("ts_issecurityinspectionsite") != null &&
                            operation.GetAttributeValue<OptionSetValue>("ts_issecurityinspectionsite").Value == 717750001)
                        {
                            si = true;
                        }
                    }
                    logMessage = String.Format("Operation Name: {0}, VSI: {1}, SI: {2}, Risk score: {3}", operation.Attributes["ovs_name"], vsi, si, riskScore);
                    logger.Info(logMessage);
                    stringBuilder.AppendLine(logMessage);

                    if (riskThresholds.Entities.Count > 0 && riskScore > 0)
                    {
                        var riskThreshold = riskThresholds.Entities.Where(a => a.GetAttributeValue<EntityReference>("ts_operationtype").Id == operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Id
                        && a.GetAttributeValue<int>("ts_riskscoreminimum") <= riskScore && a.GetAttributeValue<int>("ts_riskscoremaximum") >= riskScore).FirstOrDefault();

                        if (riskThreshold != null)
                        {
                            //Internal: Cycle length in years
                            //Frequency: How many work orders per cycle
                            var interval = riskThreshold.GetAttributeValue<int>("ts_interval");
                            var frequency = riskThreshold.GetAttributeValue<int>("ts_frequency");
                            if (frequency > 0 && interval > 0)
                            {
                                fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
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
                                        <condition attribute='ovs_operationid' operator='eq' uitype='ovs_operation' value='" + operation.Id + @"' />
                                        <condition attribute='statecode' value='0' operator='eq'/>
                                        </filter>
                                        <link-entity name='tc_tcfiscalyear' from='tc_tcfiscalyearid' to='ovs_fiscalyear' link-type='inner' alias='ag'>
                                        <filter type='and'>
                                        <condition attribute='tc_fiscalstart' operator='last-x-years' value='" + interval + @"'/>
                                        </filter>
                                        </link-entity>
                                        </entity>
                                        </fetch>";

                                EntityCollection workorders = svc.RetrieveMultiple(new FetchExpression(fetchQuery));

                                logMessage = String.Format("In total there are {0} work order(s) for operation id {1}, name {2}.", workorders.Entities.Count, operation.Id, operation.Attributes["ovs_name"]);
                                logger.Info(logMessage);
                                stringBuilder.AppendLine(logMessage);

                                //Create new work order if exists work orders within passed interval years are less then frequency
                                if (workorders.Entities != null && workorders.Entities.Count < frequency)
                                {
                                    for (int i = 0; i < frequency - workorders.Entities.Count; i++)
                                    {
                                        //Check if we are using TDG - If we are then check the Dangerous Goods
                                        var createWO = true;

                                        if (altFlag.Equals("TDG", StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            // Undecided or Schedule 1 Dangerous Goods or Non-Schedule 1 Dangerous Goods
                                            if (operation.GetAttributeValue<OptionSetValue>("ts_typeofdangerousgoods") == null ||
                                                operation.GetAttributeValue<OptionSetValue>("ts_typeofdangerousgoods").Value == 717750001 ||
                                                operation.GetAttributeValue<OptionSetValue>("ts_typeofdangerousgoods").Value == 717750002)
                                            {
                                                createWO = true;
                                            }
                                            // No Dangerous Goods
                                            else if(operation.GetAttributeValue<OptionSetValue>("ts_typeofdangerousgoods").Value == 717750000)
                                            {
                                                createWO = false;
                                            }
                                        }

                                        //Check if we are using PAX - If we are then check what kind of station is identified
                                        bool usingPAX = altFlag.Equals("PAX", StringComparison.InvariantCultureIgnoreCase);
                                        bool createSecondWorkOrder = false;

                                        if (usingPAX)
                                        {
                                            createWO = true;
                                        }

                                        if (createWO)
                                        {
                                            logMessage = String.Format("Create planning work order for operation: {0}, stake holder: {1}, type: {2}, site {3}, region {4} ",
                                            operation.Attributes["ovs_name"],
                                            operation.GetAttributeValue<EntityReference>("ts_stakeholder").Name,
                                            operation.GetAttributeValue<EntityReference>("ovs_operationtypeid").Name,
                                            operation.GetAttributeValue<EntityReference>("ts_site").Name,
                                            ((EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value).Name);

                                            logger.Info(logMessage);
                                            stringBuilder.AppendLine(logMessage);

                                            Entity workOrder = new Entity("msdyn_workorder");
                                            workOrder["msdyn_serviceaccount"] = operation.GetAttributeValue<EntityReference>("ts_stakeholder");
                                            workOrder["ovs_operationid"] = operation.ToEntityReference();
                                            workOrder["ovs_operationtypeid"] = operation.GetAttributeValue<EntityReference>("ovs_operationtypeid");
                                            workOrder["ts_site"] = operation.GetAttributeValue<EntityReference>("ts_site");
                                            workOrder["ts_region"] = (EntityReference)operation.GetAttributeValue<AliasedValue>("siteregion.ts_region").Value;
                                            if ((vsi || si) && i % 2 > 0)
                                            {
                                                //Use alt id
                                                workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(altIncidentTypeId));
                                            }
                                            else
                                            {
                                                workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(incidentTypeId));
                                            }

                                            if (usingPAX)
                                            {
                                                // Undecided or Central Station
                                                if (operation.GetAttributeValue<OptionSetValue>("ts_stationtype") == null ||
                                                    operation.GetAttributeValue<OptionSetValue>("ts_stationtype").Value == 717750000)
                                                {
                                                    //Make the Work Order Activity Type a 'Central Station Site Inspection'
                                                    workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(incidentTypeId));

                                                    //If the Operation is an SI
                                                    if (si)
                                                    {
                                                        //Set the flag to create a second Work Order with Activity Type 'Oversight of Security Inspection (PAX)'
                                                        createSecondWorkOrder = true;
                                                    }
                                                }
                                                // Transit/Unstaffed
                                                else if (operation.GetAttributeValue<OptionSetValue>("ts_stationtype").Value == 717750001)
                                                {
                                                    //Don't create a second Work Order
                                                    createSecondWorkOrder = false;

                                                    //Make the Work Order Activity Type 'Public Transit/Unstaffed Site Inspection'
                                                    workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(altSecondIncidentTypeId));
                                                }
                                            }

                                            var tradeName = Utilities.GetDefaultTradeNameByStakeHolder(svc, operation.GetAttributeValue<EntityReference>("ts_stakeholder").Id.ToString());
                                            if (tradeName != null)
                                            {
                                                workOrder["ts_tradenameid"] = tradeName;
                                            }

                                            // Set the Category to 'Planned'
                                            workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(Environment.GetEnvironmentVariable("ROM_Category_PlannedId")));  //Draft

                                            //Set the state to 'Draft'
                                            workOrder["ts_state"] = new OptionSetValue(Convert.ToInt32(717750000));

                                            //Set the Origin to next years forecast
                                            workOrder["ts_origin"] = String.Format("Forecast {0}/{1}", (DateTime.Now.AddYears(1)).ToString("yyyy"), (DateTime.Now.AddYears(2)).ToString("yy"));

                                            Guid workOrderId = svc.Create(workOrder);

                                            logMessage = String.Format("New Work Order Id: {0}", workOrderId);
                                            logger.Info(logMessage);
                                            stringBuilder.AppendLine(logMessage);

                                            //Check if a second Work Order is required
                                            if (createSecondWorkOrder)
                                            {
                                                //Make the Work Order Activity Type 'Oversight of Security Inspection (PAX)'
                                                workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid(altIncidentTypeId));
                                                workOrderId = svc.Create(workOrder);

                                                logMessage = String.Format("New Work Order Id: {0}", workOrderId);
                                                logger.Info(logMessage);
                                                stringBuilder.AppendLine(logMessage);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //Enable the follow line if want to delete generated test work orders.
                {
                    //fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //                    <entity name='msdyn_workorder'>
                    //                    <attribute name='msdyn_name'/>
                    //                    <attribute name='msdyn_workorderid'/>
                    //                    <filter type='and'>
                    //                    <filter type='or'>
                    //                    <condition attribute='msdyn_primaryincidenttype' operator='eq' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/>
                    //                    <condition attribute='msdyn_primaryincidenttype' operator='eq' uitype='msdyn_incidenttype' value='" + altIncidentTypeId + @"'/>
                    //                    </filter>                                        
                    //                    <condition attribute='statecode' value='0' operator='eq'/>
                    //                    </filter>
                    //                    <link-entity name='tc_tcfiscalyear' from='tc_tcfiscalyearid' to='ovs_fiscalyear' link-type='inner' alias='ag'>
                    //                    <filter type='and'>
                    //                    <condition attribute='tc_fiscalstart' operator='last-x-years' value='5'/>
                    //                    </filter>
                    //                    </link-entity>
                    //                    </entity>
                    //                    </fetch>";

                    //EntityCollection workordersToKeep = svc.RetrieveMultiple(new FetchExpression(fetchQuery));

                    //logMessage = Utilities.DeleteWorkOrders(svc, incidentTypeId, workordersToKeep);
                    //stringBuilder.AppendLine(logMessage);
                    //logMessage = "";
                    //logMessage = Utilities.DeleteWorkOrders(svc, altIncidentTypeId, workordersToKeep);
                    //stringBuilder.AppendLine(logMessage);
                }
            }
            catch (Exception ex)
            {
                stringBuilder.AppendLine(ex.Message);
                logger.Error(ex.Message);
            }

            return stringBuilder.ToString();
        }
    }
}
