using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

//Modified version of Aiden Kaskela's Run Workflow on Results.
//Runs an action, passing the calling System Job (asyncoperation)'s regarding field as an input parameter. 
//The action called MUST have an entityreference input parameter named "ContextEntity" matching the entity type of the target entity of the workflow triggering the query.
//workflow and actions are both the same entity, with a Category Option value indicating it is an 3 action. 
 
//Work starated on 9/6/19 by Ryan C. Perry.


namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class QueryRunActionWithContext : ContributingClasses.WorkflowQueryBase
    {
        [RequiredArgument]
        [Input("Action to Run")]
        [ReferenceTarget("workflow")]
        public InArgument<EntityReference> Workflow { get; set; }

        [Output("Number of Workflows Started")]
        public OutArgument<int> NumberOfWorkflowsStarted { get; set; }

        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            //Retrieve reference to triggering entity from passed in context. 
            EntityReference contextEntityRef = (EntityReference)workflowContext.InputParameters["Target"];

            QueryResult result = ExecuteQueryForRecords(context);
            Entity workflow = service.Retrieve("workflow", this.Workflow.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("primaryentity"));

            if (workflow.GetAttributeValue<string>("primaryentity").ToLower() != result.EntityName)
            {
                throw new ArgumentException($"Workflow entity ({workflow.GetAttributeValue<string>("primaryentity")} does not match query entity ({result.EntityName})");
            }

            int numberStarted = 0;
            foreach (Guid id in result.RecordIds)
            {
                try
                {
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest() { EntityId = id, WorkflowId = workflow.Id };
                    
                    // Add the contextEntityReference as a parameter. 
                    request.Parameters.Add("ContextEntity", contextEntityRef);

                    ExecuteWorkflowResponse response = service.Execute(request) as ExecuteWorkflowResponse;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error initiating workflow on record with ID = {id}; {ex.Message}");
                }
                numberStarted++;
            }
            this.NumberOfWorkflowsStarted.Set(context, numberStarted);
        
        }
    }
}