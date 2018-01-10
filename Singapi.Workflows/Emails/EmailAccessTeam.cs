using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;

namespace Singapi.Workflows.Emails
{
    using Common;

    public class EmailAccessTeam : CodeActivity
    {
        protected override void Execute(CodeActivityContext executionContext)
        {
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            ServiceFactory services = new ServiceFactory
            {
                Log = new StringBuilder(),
                Service = serviceFactory.CreateOrganizationService(context.UserId)
            };

            try
            {
                Guid? entityId = (context.InputParameters["Target"] as Entity)?.Id;
                if (entityId == null)
                {
                    entityId = (context.InputParameters["Target"] as EntityReference)?.Id;
                    if (entityId == null)
                    {
                        throw new InvalidOperationException("The Target parameter don't have to be null");
                    }
                }

                var sendEmail = Send.Get(executionContext);
                var teamTemplateIdsString = TeamTemplateIds.Get(executionContext);
                var emailReference = Email.Get(executionContext);

                UpdateTo(services, emailReference.Id, entityId.Value, teamTemplateIdsString);

                //Send
                if (sendEmail)
                {
                    SendEmail(services, emailReference.Id);
                }
            }
            catch (Exception ex)
            {
                services.Log.AppendLine($"Exception: {ex.ToString()}");
                throw new InvalidPluginExecutionException(ex.Message);
            }
            finally
            {
                executionContext.GetExtension<ITracingService>().Trace(services.Log.ToString());
            }
        }

        private void SendEmail(ServiceFactory services, Guid emailId)
        {
            var request = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = string.Empty,
                IssueSend = true
            };

            services.Service.Execute(request);
        }

        private void UpdateTo(ServiceFactory services, Guid emailId, Guid regardingObjectId, string teamTemplateIdsString)
        {
            services.Log.AppendLine("UpdateTo");

            var toList = new List<Entity>();

            //get email entity
            var email = services.Service.Retrieve("email", emailId, new ColumnSet("to"));
            if (email == null)
            {
                throw new InvalidOperationException($"The Email with Id ({emailId}) isn't found");
            }

            //Add already added recipients
            toList.AddRange(email.GetAttributeValue<EntityCollection>("to").Entities);

            IEnumerable<Guid> teamTemplateIds = GetTeamTemplateIds(services, teamTemplateIdsString);
            EntityCollection teamMembers = GetAccessTeamMembers(services, regardingObjectId, teamTemplateIds);

            PopulateUsers(services, teamMembers, ref toList);

            email["to"] = toList.ToArray();

            //send changes
            services.Service.Update(email);
        }

        private IEnumerable<Guid> GetTeamTemplateIds(ServiceFactory services, string teamTemplateIdsString)
        {
            services.Log.AppendLine("GetTeamTemplateIds");

            var result = new List<Guid>();

            foreach (var teamTemplateIdString in teamTemplateIdsString.Split(','))
            {
                Guid readlId;
                if (Guid.TryParse(teamTemplateIdString, out readlId))
                {
                    result.Add(readlId);
                }
                else
                {
                    services.Log.AppendLine($"Warn: The Team Template Id ({teamTemplateIdString}) cannot be conveted to Guid data type ");
                }
            }

            return result;
        }

        private EntityCollection GetAccessTeamMembers(ServiceFactory services, Guid objectId, IEnumerable<Guid> teamTemplateIds)
        {
            services.Log.AppendLine("GetAccessTeamMembers");
            services.Log.AppendLine($"teamTemplateIds = {teamTemplateIds.Select(s => s.ToString()).Aggregate((l, r) => $"{l},{r}")}");

            var query = new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet("systemuserid")
            };

            var teamLink = query.AddLink("team", "teamid", "teamid");
            teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 1);
            teamLink.LinkCriteria.AddCondition("regardingobjectid", ConditionOperator.Equal, objectId);

            //add team templates in condition
            var teamCondition = new ConditionExpression("teamtemplateid", ConditionOperator.In);
            foreach (var teamTemplateId in teamTemplateIds)
            {
                teamCondition.Values.Add(teamTemplateId);
            }
            teamLink.LinkCriteria.AddCondition(teamCondition);

            return services.Service.RetrieveMultiple(query);
        }

        private void PopulateUsers(ServiceFactory services, EntityCollection teamMembers, ref List<Entity> toList)
        {
            services.Log.AppendLine("PopulateUsers");

            var addedUsers = new Dictionary<Guid, bool>();
            foreach (Entity entity in teamMembers.Entities)
            {
                var userId = entity.GetAttributeValue<Guid>("systemuserid");
                if (!addedUsers.ContainsKey(userId))
                {
                    addedUsers.Add(userId, true);
                }
            }

            foreach (Guid userId in addedUsers.Keys)
            {
                Entity user = services.Service.Retrieve("systemuser", userId, new ColumnSet("internalemailaddress"));

                if (string.IsNullOrEmpty(user.GetAttributeValue<string>("internalemailaddress")))
                {
                    services.Log.AppendLine($"Warn: The System User with Id ({userId}) internalemailaddress attribute is not populated");
                    continue;
                }

                var activityParty = new Entity("activityparty");
                activityParty["partyid"] = new EntityReference("systemuser", userId);
                toList.Add(activityParty);
            }
        }

        [RequiredArgument]
        [Input("Email")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

        [RequiredArgument]
        [Input("Team Template Ids (comma-separated)")]
        public InArgument<string> TeamTemplateIds { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Send?")]
        public InArgument<bool> Send { get; set; }
    }
}
