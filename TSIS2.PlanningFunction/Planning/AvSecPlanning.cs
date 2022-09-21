using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Azure.WebJobs.Host;
using System.Text;
using System.Collections.Generic;
using System.Xml;
using System.IO;
using Microsoft.Xrm.Sdk.Messages;

namespace TSIS2.PlanningFunction
{
    public class AvSecPlanning
    {
        public string GenerateWorkOrders(CrmServiceClient svc, TraceWriter logger)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                logger.Info(String.Format("Start processing AcSecPlanning."));
                string fetchXmlFrequencies = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='ts_recurrencefrequencies'>
                                        <attribute name='ts_recurrencefrequenciesid' />
                                        <attribute name='ts_name' />
                                        <attribute name='ts_class1interval' />
                                        <attribute name='ts_class1frequency' />
                                        <attribute name='ts_class2and3highriskinterval' />
                                        <attribute name='ts_class2and3highriskfrequency' />
                                        <attribute name='ts_class2and3lowriskinterval' />
                                        <attribute name='ts_class2and3lowriskfrequency' />
                                        <order attribute='ts_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                List<Entity> recurrencefrequencies = svc.RetrieveMultiple(new FetchExpression(fetchXmlFrequencies)).Entities.ToList();

                string fetchXmlSiteVisit = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='ts_sitevisit'>
                                        <attribute name='ts_sitevisitid' />
                                        <attribute name='ts_name' />
                                        <attribute name='ts_status' />
                                        <attribute name='ts_functionallocation' />
                                        <order attribute='ts_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='ts_fiscalyear' operator='not-null' />
                                        </filter>
                                        <link-entity name='tc_tcfiscalyear' from='tc_tcfiscalyearid' to='ts_fiscalyear' link-type='inner' alias='al'>
                                          <filter type='and'>
                                            <condition attribute='tc_fiscalend' operator='on-or-after' value='" + GetCurrentFiscalYearEndDate().AddYears(-1).AddDays(1).ToString("yyyy-MM-dd") + @"' />
                                            <condition attribute='tc_fiscalend' operator='on-or-before' value='" + GetCurrentFiscalYearEndDate().ToString("yyyy-MM-dd") + @"' />
                                          </filter>
                                        </link-entity>
                                      </entity>
                                    </fetch>";
                List<Entity> siteVisits = svc.RetrieveMultiple(new FetchExpression(fetchXmlSiteVisit)).Entities.ToList();

                string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='ts_operationactivity'>
                                    <attribute name='ts_operationactivityid' />
                                    <attribute name='ts_name' />
                                    <attribute name='ts_site' />
                                    <attribute name='ts_operation' />
                                    <attribute name='ts_duedate' />
                                    <attribute name='ts_activity' />
                                    <attribute name='ts_stakeholder' />
                                    <order attribute='ts_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='ts_operationalstatus' operator='eq' value='717750000' />
                                    </filter>
                                    <link-entity name='msdyn_functionallocation' from='msdyn_functionallocationid' to='ts_site' link-type='inner' alias='oaf'>
                                      <attribute name='ts_sitetype' />
                                      <attribute name='ts_class' />
                                      <attribute name='ts_region' />
                                      <attribute name='ts_riskscore' />
                                      <filter type='and'>
                                        <condition attribute='ts_class' operator='in'>
                                          <value>717750001</value>
                                          <value>717750002</value>
                                          <value>717750003</value>
                                        </condition>
                                        <condition attribute='ts_sitetype' operator='eq' uiname='Aerodrome' uitype='ovs_sitetype' value='{99DA31E7-7D78-EB11-A812-0022486D697D}' />
                                      </filter>
                                    </link-entity>
                                    <link-entity name='ovs_operation' from='ovs_operationid' to='ts_operation' link-type='inner' alias='oao'>
                                      <attribute name='ovs_operationtypeid' />
                                      <attribute name='ts_operationalstatus' />
                                      <filter type='and'>
                                        <condition attribute='ts_operationalstatus' operator='eq' value='717750000' />
                                      </filter>
                                    </link-entity>
                                    <link-entity name='tc_tcfiscalquarter' from='tc_tcfiscalquarterid' to='ts_duedate' visible='false' link-type='outer' alias='oat'>
                                      <attribute name='tc_tcfiscalyearid' />
                                      <attribute name='tc_quarterstart' />
                                      <attribute name='tc_quarterend' />
                                    </link-entity>
                                    <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='ts_activity' visible='false' link-type='outer' alias='oai'>
                                      <attribute name='ts_riskscore' />
                                    </link-entity>
                                  </entity>
                                </fetch>";

                // Define the fetch attributes.
                // Set the number of records per page to retrieve.
                int fetchCount = 1000;
                // Initialize the page number.
                int pageNumber = 1;
                // Initialize the number of records.
                int recordCount = 0;
                // Specify the current paging cookie. For retrieving the first page, 
                // pagingCookie should be null.
                string pagingCookie = null;

                while (true)
                {
                    // Build fetchXml string with the placeholders.
                    string xml = CreateXml(fetchXml, pagingCookie, pageNumber, fetchCount);

                    // Excute the fetch query and get the xml result.
                    RetrieveMultipleRequest fetchRequest = new RetrieveMultipleRequest
                    {
                        Query = new FetchExpression(xml)
                    };

                    EntityCollection operationActivities = ((RetrieveMultipleResponse)svc.Execute(fetchRequest)).EntityCollection;
                    recordCount += operationActivities.Entities.Count;
                    foreach (var operationActivity in operationActivities.Entities)
                    {
                        var riskscoreId = ((EntityReference)operationActivity.GetAttributeValue<AliasedValue>("oai.ts_riskscore")?.Value)?.Id;
                        Entity riskscore = null;
                        //Internal: Cycle length in years
                        //Frequency: How many work orders per cycle
                        int frequency = 0;
                        int interval = 0;
                        if (riskscoreId != null)
                        {
                            riskscore = recurrencefrequencies.Where(a => a.Id == riskscoreId).FirstOrDefault();
                        }
                        bool isDue = false;
                        bool isUnplanned = false;
                        if (operationActivity.GetAttributeValue<EntityReference>("ts_duedate") != null)
                        {
                            if (operationActivity.GetAttributeValue<AliasedValue>("oat.tc_quarterend") != null)
                            {
                                var dueDate = (DateTime)operationActivity.GetAttributeValue<AliasedValue>("oat.tc_quarterend").Value;
                                var planningDate = GetCurrentFiscalYearEndDate(); 

                                if (dueDate.Date <= planningDate.Date)
                                {
                                    isDue = true;
                                }
                            }
                        }

                        if (((OptionSetValue)operationActivity.GetAttributeValue<AliasedValue>("oaf.ts_class").Value).Value == 717750001) //Class 1
                        {
                            if (riskscore != null)
                            {
                                frequency = riskscore.GetAttributeValue<int>("ts_class1frequency");
                                interval = riskscore.GetAttributeValue<int>("ts_class1interval");
                            }
                        }
                        else //Class 2 & 3
                        {
                            if (riskscore != null)
                            {
                                int siteRiskscore = 0;
                                if (operationActivity.GetAttributeValue<AliasedValue>("oaf.ts_riskscore") != null)
                                {
                                    siteRiskscore = (int)operationActivity.GetAttributeValue<AliasedValue>("oaf.ts_riskscore").Value;
                                }
                                if (siteRiskscore > 5)
                                {
                                    frequency = riskscore.GetAttributeValue<int>("ts_class2and3highriskfrequency");
                                    interval = riskscore.GetAttributeValue<int>("ts_class2and3highriskinterval");
                                }
                                else
                                {
                                    frequency = riskscore.GetAttributeValue<int>("ts_class2and3lowriskfrequency");
                                    interval = riskscore.GetAttributeValue<int>("ts_class2and3lowriskinterval");
                                }    
                            }

                            //If Has Site Visit
                            var siteVisit = siteVisits.Where(a => a.GetAttributeValue<EntityReference>("ts_functionallocation").Id == operationActivity.GetAttributeValue<EntityReference>("ts_site").Id).FirstOrDefault();
                            if (siteVisit != null)
                            {
                                if (!isDue)
                                {
                                    isUnplanned = true;
                                }
                            }
                        }

                        if (isDue || isUnplanned)
                        {
                            int recordToCreated = (int)Math.Ceiling((decimal)frequency / ((interval == 0)?1: interval));
                            for (var i = 0; i < recordToCreated; i++)
                            {
                                sb.AppendLine(String.Format("Create planning work order for operation: {0}, stake holder: {1}, type: {2}, site {3}, region {4}, activity {5}",
                                             operationActivity.GetAttributeValue<EntityReference>("ts_operation").Name,
                                             operationActivity.GetAttributeValue<EntityReference>("ts_stakeholder").Name,
                                            ((EntityReference)operationActivity.GetAttributeValue<AliasedValue>("oao.ovs_operationtypeid").Value).Name,
                                            operationActivity.GetAttributeValue<EntityReference>("ts_site").Name,
                                            ((EntityReference)operationActivity.GetAttributeValue<AliasedValue>("oaf.ts_region").Value).Name,
                                            operationActivity.GetAttributeValue<EntityReference>("ts_activity").Name));

                                Entity workOrder = new Entity("msdyn_workorder");
                                workOrder["msdyn_serviceaccount"] = operationActivity.GetAttributeValue<EntityReference>("ts_stakeholder");
                                workOrder["ovs_operationid"] = operationActivity.GetAttributeValue<EntityReference>("ts_operation");
                                workOrder["ovs_operationtypeid"] = (EntityReference)operationActivity.GetAttributeValue<AliasedValue>("oao.ovs_operationtypeid").Value;
                                workOrder["ts_site"] = operationActivity.GetAttributeValue<EntityReference>("ts_site");
                                workOrder["ts_region"] = (EntityReference)operationActivity.GetAttributeValue<AliasedValue>("oaf.ts_region").Value;
                                workOrder["msdyn_primaryincidenttype"] = operationActivity.GetAttributeValue<EntityReference>("ts_activity");
                                var tradeName = Utilities.GetDefaultTradeNameByStakeHolder(svc, operationActivity.GetAttributeValue<EntityReference>("ts_stakeholder").Id.ToString());
                                if (tradeName != null)
                                {
                                    workOrder["ts_tradenameid"] = tradeName;
                                }
                                if (isUnplanned)
                                {
                                    workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(Environment.GetEnvironmentVariable("ROM_Category_UnplannedId", EnvironmentVariableTarget.Process)));  //UnPlanned
                                }
                                else
                                {
                                    workOrder["ovs_rational"] = new EntityReference("ovs_tyrational", new Guid(Environment.GetEnvironmentVariable("ROM_Category_PlannedId", EnvironmentVariableTarget.Process)));  //Planned
                                }
                                workOrder["ts_state"] = new OptionSetValue(Convert.ToInt32(717750000));   //Draft
                                workOrder["ts_origin"] = String.Format("Forecast {0}/{1}", (DateTime.Now.AddYears(1)).ToString("yyyy"), (DateTime.Now.AddYears(2)).ToString("yy"));
                                Guid workOrderId = svc.Create(workOrder);
                                sb.AppendLine(String.Format("New Work Order Id: {0}", workOrderId));
                            }
                        }
                    }

                    // Check for morerecords, if it returns 1.
                    if (operationActivities.MoreRecords)
                    {
                        // Increment the page number to retrieve the next page.
                        pageNumber++;

                        // Set the paging cookie to the paging cookie returned from current results.                            
                        pagingCookie = operationActivities.PagingCookie;
                    }
                    else
                    {
                        // If no more records in the result nodes, exit the loop.
                        break;
                    }
                }

                sb.AppendLine(String.Format("In total of {0} related operation activity records", recordCount));
            }
            catch (Exception ex)
            {
                sb.AppendLine(ex.Message);
                logger.Error(ex.Message);
            }
            return sb.ToString();
        }

        public string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }

        public DateTime GetCurrentFiscalYearEndDate()
        {
            return new DateTime(DateTime.Now.Year + ((DateTime.Now.Month > 3) ? 1 : 0), 3, 31);
        }
    }
}
