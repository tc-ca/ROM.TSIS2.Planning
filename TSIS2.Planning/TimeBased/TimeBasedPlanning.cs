using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace TSIS2.Planning
{
    public class TimeBasedPlanning
    {
        public void GenerateWorkOrderByIncidentType(CrmServiceClient svc, string incidentTypeId)
        {
            string fetchquery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name'/>
                        <attribute name='createdon'/>
                        <attribute name='msdyn_serviceaccount'/>
                        <attribute name='msdyn_workorderid'/>
                        <attribute name='msdyn_functionallocation'/>
                        <order attribute='msdyn_name' descending='false'/>
                        <filter type='and'>
                        <condition attribute='msdyn_primaryincidenttype' operator='eq' uiname='Security Plan Review (TDG)' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/>
                        </filter>
                        <link-entity name='tc_tcfiscalyear' from='tc_tcfiscalyearid' to='ovs_fiscalyear' link-type='inner' alias='ag'>
                        <filter type='and'>
                        <condition attribute='tc_fiscalstart' operator='last-x-years' value='5'/>
                        </filter>
                        </link-entity>
                        </entity>
                        </fetch>";

            EntityCollection workorders = svc.RetrieveMultiple(new FetchExpression(fetchquery));
            Console.WriteLine("In total there are {0} Security Plan Review (TDG) work order(s).", workorders.Entities.Count);
            foreach (var c in workorders.Entities)
            {
                Console.WriteLine("Work Order Name: {0}", c.Attributes["msdyn_name"]);
            }

            fetchquery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='ovs_operationtype'>
                        <attribute name='ovs_operationtypeid'/>
                        <attribute name='ovs_name'/>
                        <attribute name='createdon'/>
                        <order attribute='ovs_name' descending='false'/>
                        <link-entity name='ts_ovs_operationtypes_msdyn_incidenttypes' from='ovs_operationtypeid' to='ovs_operationtypeid' visible='false' intersect='true'>
                        <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' alias='ab'>
                        <filter type='and'>
                        <condition attribute='msdyn_incidenttypeid' operator='eq' uiname='Security Plan Review (TDG)' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"'/>
                        </filter>
                        </link-entity>
                        </link-entity>
                        </entity>
                        </fetch>";

            EntityCollection operationtypes = svc.RetrieveMultiple(new FetchExpression(fetchquery));
            Console.WriteLine("In total there are {0} Operation Types that related to Security Plan Review (TDG).", operationtypes.Entities.Count);
            foreach (var c in operationtypes.Entities)
            {
                Console.WriteLine("Operation type Name: {0}", c.Attributes["ovs_name"]);
            }

            //All operations related to incident type Security Plan Review (TDG)
            fetchquery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='ovs_operation'>
                    <attribute name='ovs_name' />
                    <attribute name='createdon' />
                    <attribute name='ovs_operationtypeid' />
                    <attribute name='ovs_operationid' />
                    <attribute name='ts_stakeholder' />
                    <order attribute='ovs_name' descending='false' />
                    <link-entity name='ovs_operationtype' from='ovs_operationtypeid' to='ovs_operationtypeid' link-type='inner' alias='ac'>
                      <link-entity name='ts_ovs_operationtypes_msdyn_incidenttypes' from='ovs_operationtypeid' to='ovs_operationtypeid' visible='false' intersect='true'>
                        <link-entity name='msdyn_incidenttype' from='msdyn_incidenttypeid' to='msdyn_incidenttypeid' alias='ad'>
                          <filter type='and'>
                            <condition attribute='msdyn_incidenttypeid' operator='eq' uiname='Security Plan Review (TDG)' uitype='msdyn_incidenttype' value='" + incidentTypeId + @"' />
                          </filter>
                        </link-entity>
                      </link-entity>
                    </link-entity>
                  </entity>
                </fetch>";

            EntityCollection operations = svc.RetrieveMultiple(new FetchExpression(fetchquery));
            Console.WriteLine("In total there are {0} Operations that related to Security Plan Review (TDG).", operations.Entities.Count);
            foreach (var c in operations.Entities)
            {
                Console.WriteLine("Operation Name: {0}, Stake holder: {1},  Operation Type: {2} ", c.Attributes["ovs_name"], c.GetAttributeValue<EntityReference>("ts_stakeholder").Name, c.GetAttributeValue<EntityReference>("ovs_operationtypeid").Name);

            }

            //Entity workOrder = new Entity("msdyn_workorder");
            //workOrder["msdyn_name"] = "300-345679";
            //workOrder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("bd01824a-dbeb-eb11-bacb-000d3af4fbec"));
            //workOrder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid("21c59aa0-511a-ec11-b6e7-000d3a09ce95"));
            //Guid workOrderId = svc.Create(workOrder);
            //Console.WriteLine("New Work Order Id: {0}", workOrderId);
        }
    }
}
