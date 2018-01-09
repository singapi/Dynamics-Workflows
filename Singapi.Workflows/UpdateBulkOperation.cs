namespace Singapi.Workflows
{
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;
    using System;
    using System.Activities;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Linq;
    using System.Xml.Linq;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Common;
    using Microsoft.Xrm.Sdk.Client;

    public class UpdateBulkOperation : CodeActivity
    {
        // Define the fetch attributes.
        // Set the number of records per page to retrieve.
        const int FetchCount = 5000;

        private IOrganizationService _service;

        [RequiredArgument]
        [Input("FetchXml (entities to update)")]
        public InArgument<string> FetchXml { get; set; }

        //[Input("FetchXml Parameters")]
        //public InArgument<string> FetchXmlParameters { get; set; }

        [RequiredArgument]
        [Input("Attributes (JSON)")]
        public InArgument<string> Attributes { get; set; }

        [Input("Attributes Parameters")]
        public InArgument<string> AttributesParameters { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            StringBuilder sb = new StringBuilder();
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            _service = serviceFactory.CreateOrganizationService(context.UserId);

            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(_service))
            {
                try
                {
                    var fetchXml = FetchXml.Get(executionContext);
                    //var fetchXmlParameters = FetchXmlParameters.Get(executionContext);
                    var attributesJsonString = Attributes.Get(executionContext);
                    var attributesParametersString = AttributesParameters.Get(executionContext);

                    var entity = (Entity)context.InputParameters["Target"];

                    #region prepare parameters

                    var functions = new Expressive.Functions.IFunction[] { new LookupFieldFunction(orgContext), new FieldFunction(orgContext) };

                    var attributeParameters = attributesParametersString.Split(',').Select(s => s.Trim()).ToArray();
                    for (var i = 0; i < attributeParameters.Length; i++)
                    {
                        var parameterExpression = Helper.Create(attributeParameters[i], functions);
                        attributesJsonString = attributesJsonString.Replace("{{" + i + "}}", parameterExpression.Evaluate(new Dictionary<string, object> { ["entity"] = entity }).ToString());
                    }

                    #endregion

                    //var expression = Helper.Create(attributeParameters[i], functions);

                    //expression.Evaluate(new Dictionary<string, object> { ["entity"] = entity });
                    

                    var xdoc = XDocument.Parse(fetchXml);
                    var attributesMetadata = GetAttribuitesMetadata(xdoc);
                    var crmAttributeConverter = new CrmAttributeConverter(attributesMetadata);
                    var attributes = JsonConvert.DeserializeObject<Dictionary<string, object>>(attributesJsonString, crmAttributeConverter);

                    var entities = RetrieveEntities(xdoc);
                    var attributesToUpdate = new Dictionary<string, object>();

                    foreach (var attributeItem in attributes)
                    {
                        var attr = attributesMetadata.SingleOrDefault(a => a.LogicalName.Equals(attributeItem.Key, StringComparison.OrdinalIgnoreCase));
                        if (attr == null)
                        {
                            //todo: write message Attribute is not found
                            sb.AppendLine($"Attribute {attributeItem.Key} isn't found");
                            continue;
                        }
                        //if (!attr.IsValidForUpdate.GetValueOrDefault(false))
                        //{
                        //    sb.AppendLine($"Isn't allowing to update {attr.LogicalName} attribute");
                        //    continue;
                        //}
                        attributesToUpdate.Add(attr.LogicalName, attributeItem.Value);
                    }

                    var requestWithResults = new ExecuteMultipleRequest()
                    {
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = true
                        },
                        Requests = CreateBulkUpdateRequests(entities, attributes)
                    };

                    var response = (ExecuteMultipleResponse)_service.Execute(requestWithResults);
                }
                catch (Exception e)
                {
                    throw new InvalidPluginExecutionException(e.Message);
                }
            }
        }

        private AttributeMetadata[] GetAttribuitesMetadata(XDocument xdoc)
        {
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = xdoc.Root.Element("entity").Attribute("name").Value
            };

            var response = (RetrieveEntityResponse)_service.Execute(retrieveEntityRequest);
            return response.EntityMetadata.Attributes;
        }

        private IEnumerable<Entity> RetrieveEntities(XDocument xdoc)
        {
            int pageNumber = 1;
            string pagingCookie = null;
            List<Entity> returnCollections = new List<Entity>();

            while (true)
            {
                string xml = ModifyXml(xdoc, pagingCookie, pageNumber, FetchCount);

                //ExtractAttribute()

                var fetchRequest = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                EntityCollection returnCollection = _service.RetrieveMultiple(new FetchExpression(xml));
                returnCollections.AddRange(returnCollection.Entities);
                if (returnCollection.MoreRecords)
                {
                    pageNumber++;
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }

            return returnCollections;
        }

        private OrganizationRequestCollection CreateBulkUpdateRequests(IEnumerable<Entity> entities, IDictionary<string, object> attributesToUpdate)
        {
            var requestCollection = new OrganizationRequestCollection();

            foreach (var entity in entities)
            {
                foreach (var attributeItem in attributesToUpdate)
                {
                    if (entity.Contains(attributeItem.Key))
                    {
                        entity.Attributes[attributeItem.Key] = attributeItem.Value;
                    }
                    else
                    {
                        entity.Attributes.Add(attributeItem);
                    }
                }
                requestCollection.Add(new UpdateRequest { Target = entity });
            }
            return requestCollection;
        }

        public string ExtractNodeValue(XmlNode parentNode, string name)
        {
            XmlNode childNode = parentNode.SelectSingleNode(name);

            if (null == childNode)
            {
                return null;
            }
            return childNode.InnerText;
        }

        public string ExtractAttribute(XmlDocument doc, string name)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;
            XmlAttribute attr = (XmlAttribute)attrs.GetNamedItem(name);
            if (null == attr)
            {
                return null;
            }
            return attr.Value;
        }

        public string ModifyXml(XDocument xdoc, string cookie, int page, int count)
        {
            if (cookie != null)
            {
                xdoc.Root.Add(new XAttribute("paging-cookie", cookie));
            }

            xdoc.Root.Add(new XAttribute("page", page));
            xdoc.Root.Add(new XAttribute("count", count));

            var sb = new StringBuilder(2048);
            var stringWriter = new StringWriter(sb);

            using (var writer = new XmlTextWriter(stringWriter))
            {
                xdoc.WriteTo(writer);
            }

            return sb.ToString();
        }

        private class CrmAttributeConverter : JsonConverter
        {
            private readonly AttributeMetadata[] _attributesMetadata;

            public CrmAttributeConverter(AttributeMetadata[] attributesMetadata)
            {
                _attributesMetadata = attributesMetadata;
            }

            public override bool CanConvert(Type objectType)
            {
                return true;
            }

            public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
            {
                //JObject jsonObject = JObject.Load(reader);
                //var properties = jsonObject.Properties().ToList();
                var result = new Dictionary<string, object>();


                if (reader.TokenType == JsonToken.StartObject)
                {
                    //var attr = _attributesMetadata.SingleOrDefault(a => a.LogicalName.Equals(reader.Value.ToString(), StringComparison.OrdinalIgnoreCase));
                    JObject item = JObject.Load(reader);

                    if (objectType.Equals(typeof(Dictionary<string, object>)))
                    { 
                        foreach (var propertyItem in item.Properties())
                        {
                            var attr = _attributesMetadata.SingleOrDefault(a => a.LogicalName.Equals(propertyItem.Name, StringComparison.OrdinalIgnoreCase));

                            switch (attr.AttributeType)
                            {
                                case AttributeTypeCode.BigInt:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<long>());
                                    break;
                                case AttributeTypeCode.Boolean:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<bool>());
                                    break;
                                case AttributeTypeCode.DateTime:
                                    result.Add(attr.LogicalName, DateTime.SpecifyKind(propertyItem.ToObject<DateTime>(), DateTimeKind.Utc));
                                    break;
                                case AttributeTypeCode.Decimal:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<decimal>());
                                    break;
                                case AttributeTypeCode.Double:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<double>());
                                    break;
                                case AttributeTypeCode.Integer:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<int>());
                                    break;
                                case AttributeTypeCode.Customer:
                                case AttributeTypeCode.Lookup:
                                    result.Add(attr.LogicalName, serializer.Deserialize<EntityReference>(propertyItem.First.CreateReader()));
                                    break;
                                case AttributeTypeCode.Money:
                                    result.Add(attr.LogicalName, serializer.Deserialize<Money>(propertyItem.First.CreateReader()));
                                    break;
                                case AttributeTypeCode.Picklist:
                                case AttributeTypeCode.State:
                                case AttributeTypeCode.Status:
                                    result.Add(attr.LogicalName, serializer.Deserialize<OptionSetValue>(propertyItem.First.CreateReader()));
                                    break;
                                case AttributeTypeCode.Memo:
                                case AttributeTypeCode.String:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<string>());
                                    break;
                                case AttributeTypeCode.Uniqueidentifier:
                                    result.Add(attr.LogicalName, propertyItem.ToObject<Guid>());
                                    break;


                            }
                        }

                        return result;
                    }
                    else if (objectType.Equals(typeof(OptionSetValue)))
                    {
                        return item.ToObject<OptionSetValue>();
                    }
                    else if (objectType.Equals(typeof(EntityReference)))
                    {
                        return item.ToObject<EntityReference>();
                    }
                    else if (objectType.Equals(typeof(Money)))
                    {
                        return item.ToObject<Money>();
                    }
                }

                throw new InvalidOperationException($"{nameof(CrmAttributeConverter)} is failed");
            }

            public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
            {
                throw new NotImplementedException("Unnecessary because CanWrite is false. The type will skip the converter.");
            }

            public override bool CanWrite { get; } = false;
        }
    }
}
