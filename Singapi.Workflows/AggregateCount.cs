using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Xml.Linq;

namespace Singapi.Workflows
{
    public class AggregateCount : CodeActivity
    {
        private IOrganizationService _service;

        [RequiredArgument]
        [Input("FetchXml (entities to update)")]
        public InArgument<string> FetchXml { get; set; }

        [Output("Count records")]
        public OutArgument<int> Count { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            StringBuilder sb = new StringBuilder();
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            _service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                var fetchXml = FetchXml.Get(executionContext);
                var xdoc = XElement.Parse(fetchXml);
                if (xdoc == null)
                {
                    throw new InvalidOperationException("FetchXml argument isn't valid");
                }
                //set aggregate attribute
                xdoc.SetAttributeValue(XName.Get(Constants.AggregateAttribute), true);
                xdoc.SetAttributeValue(XName.Get(Constants.NoLockAttribute), true);

                var entityElement = xdoc.Element("entity");
                var entityNameAttribute = entityElement.Attribute(XName.Get(Constants.NameAttribute));
                var entityName = entityNameAttribute.Value;
                var idAttributeName = $"{entityName}id";

                //remove all attributes
                entityElement.Elements(XName.Get(Constants.AttributeTagName)).Remove();

                //create aggrigate count attribute
                var aggrigateAttribute = new XElement(XName.Get(Constants.AttributeTagName));
                aggrigateAttribute.SetAttributeValue(XName.Get(Constants.NameAttribute), idAttributeName);
                aggrigateAttribute.SetAttributeValue(XName.Get(Constants.AliasAttribute), Constants.AggregateCount);
                aggrigateAttribute.SetAttributeValue(XName.Get(Constants.AggregateAttribute), Constants.AggregateCount);
                entityElement.Add(aggrigateAttribute);

                ///set output argument
                Count.Set(executionContext, GetCount(xdoc));
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private int GetCount(XContainer xdoc)
        {
            //ExtractAttribute()

            var returnCollection = _service.RetrieveMultiple(new FetchExpression(xdoc.ToString(SaveOptions.DisableFormatting)));
            var entity = returnCollection.Entities?.SingleOrDefault();
            if (entity == null)
            {
                return 0;
            }

            return entity.GetAttributeValue<int>(Constants.AggregateCount);
        }
    }
}
