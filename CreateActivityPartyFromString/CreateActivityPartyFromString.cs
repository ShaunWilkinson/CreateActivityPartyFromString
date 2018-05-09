using System;

using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace CreateActivityParty
{
    /// <summary>
    /// Allows a user to create an ActivityParty record using a string extending the functionality of the built in CRM workflows
    /// </summary>
    /// <remarks>
    /// Usage - 
    /// Create a workflow which creates the email sans values in the To field. Add this workflow step and assign a string
    /// and a reference to the email record. Then create a step to change status of the email record to sent.
    /// </remarks
    public class CreateActivityPartyFromString : CodeActivity
    {
        [RequiredArgument]
        [Input("String to Convert")]
        public InArgument<String> OriginalString { get; set; }

        [RequiredArgument]
        [Input("Email Record")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailRecord { get; set; }


        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Get the inputs
            EntityReference emailRecord = EmailRecord.Get<EntityReference>(executionContext);
            Guid emailRecordId = emailRecord.Id;
            String originalString = OriginalString.Get<String>(executionContext);

            // Create an activity party as entity collection
            Entity unresolvedParty = new Entity("activityparty");
            unresolvedParty["addressused"] = originalString;

            EntityCollection to = new EntityCollection();
            to.Entities.Add(unresolvedParty);

            // Create request for email
            ColumnSet columnsToRequest = new ColumnSet("to", "subject");
            try
            {
                Entity retrievedEmail = service.Retrieve("email", emailRecordId, columnsToRequest);
                tracingService.Trace("Retrieved Email: " + retrievedEmail["subject"].ToString());

                // Set the TO field
                retrievedEmail["to"] = to;

                tracingService.Trace("Updating Email");
                try
                {
                    // Update the record
                    service.Update(retrievedEmail);
                }
                catch (Exception err)
                {
                    tracingService.Trace("Failed to update email - " + err.Message);
                }

            }
            catch (Exception err)
            {
                tracingService.Trace("Failed to retrieve email - " + err.Message);
            }
        }
    }
}
