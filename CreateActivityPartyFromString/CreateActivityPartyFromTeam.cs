using System;

using System.Activities;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

// This namespace is found in Microsoft.Crm.Sdk.Proxy.dll assembly
// found in the SDK\bin folder.
using Microsoft.Crm.Sdk.Messages;

namespace CreateActivityParty
{
    /// <summary>
    /// Allows a user to create an ActivityParty record by entering a team
    /// </summary>
    /// <remarks>
    /// Usage - 
    /// Create a workflow which creates the email sans values in the To field. Add this workflow step and assign a team
    /// and a reference to the email record.  This workflow will send the email
    /// </remarks
    public class CreateActivityPartyFromTeam : CodeActivity
    {
        [RequiredArgument]
        [Input("Team")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> TeamRecord { get; set; }

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
            EntityReference teamRecord = TeamRecord.Get<EntityReference>(executionContext);
            Guid emailRecordId = emailRecord.Id;
            Guid teamRecordId = teamRecord.Id;

            // create the query expression
            StringBuilder linkFetch = new StringBuilder();
            linkFetch.Append("<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"true\">");
            linkFetch.Append("<entity name=\"systemuser\">");
            linkFetch.Append("<attribute name=\"fullname\"/>");
            linkFetch.Append("<attribute name=\"systemuserid\"/>");
            linkFetch.Append("<attribute name=\"internalemailaddress\"/>");
            linkFetch.Append("<link-entity name=\"teammembership\" from=\"systemuserid\" to=\"systemuserid\" visible=\"false\" intersect=\"true\">");
            linkFetch.Append("<link-entity name=\"team\" from=\"teamid\" to=\"teamid\" alias=\"aa\">");
            linkFetch.Append("<filter type=\"and\">");
            linkFetch.Append("<condition attribute=\"teamid\" operator=\"eq\" uiname=\"" + teamRecord.Name + "\" uitype=\"team\" value=\"" + teamRecordId + "\"/>");
            linkFetch.Append("</filter>");
            linkFetch.Append("</link-entity>");
            linkFetch.Append("</link-entity>");
            linkFetch.Append("</entity>");
            linkFetch.Append("</fetch>");

            // Create the retrieveRequest
            RetrieveMultipleRequest retrieveRequest = new RetrieveMultipleRequest()
            {
                Query = new FetchExpression(linkFetch.ToString())
            };
            
            // Obtain results from query
            EntityCollection returnedUsers = ((RetrieveMultipleResponse)service.Execute(retrieveRequest)).EntityCollection;

            if (returnedUsers.Entities.Count == 0)
            {
                tracingService.Trace("Couldn't locate any users in team - {0}", teamRecord.Name);
                return;
            }

            // For each of the returned users, create an activityparty record then add it to the collecton
            EntityCollection to = new EntityCollection();
            foreach(Entity user in returnedUsers.Entities)
            {
                try
                {
                    Entity activityParty = new Entity("activityparty");
                    activityParty["addressused"] = user["internalemailaddress"];
                    activityParty["partyid"] = new EntityReference("systemuser", user.Id);
                    to.Entities.Add(activityParty);
                } catch (Exception err)
                {
                    tracingService.Trace("Failed at this point: {0}", err);
                }
            }

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

                    // Create the send email request
                    SendEmailRequest sendEmailReq = new SendEmailRequest
                    {
                        EmailId = emailRecordId,
                        TrackingToken = "",
                        IssueSend = true
                    };

                    // Send the email
                    SendEmailResponse sendEmail = (SendEmailResponse)service.Execute(sendEmailReq);
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
