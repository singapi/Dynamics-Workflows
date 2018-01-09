using Expressive.Expressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Linq;
using System.Collections.Generic;
using Expressive.Functions;
using Expressive;

namespace Singapi.Workflows.Common
{
    public abstract class CrmBaseFunction : IFunction
    {
        private readonly string _name;
        protected readonly OrganizationServiceContext _orgContext;

        public CrmBaseFunction(OrganizationServiceContext context, string name)
        {
            _name = name;
            _orgContext = context;
        }

        public string Name
        {
            get
            {
                return _name;
            }
        }

        public IDictionary<string, object> Variables { get; set; }

        public abstract object Evaluate(IExpression[] parameters, ExpressiveOptions options);
    }

    public class LookupFieldFunction : CrmBaseFunction
    {
        public LookupFieldFunction(OrganizationServiceContext context)
            : base(context, "LookupField")
        {
        }

        public override object Evaluate(IExpression[] parameters, ExpressiveOptions options)
        {
            if (parameters.Length < 1)
            {
                throw new Exception($"Function {Name} must have at least one parameter");
            }
            var entity = Variables["entity"] as Entity;
            var attributePath = parameters[0].Evaluate(Variables)?.ToString();
            if (string.IsNullOrEmpty(attributePath))
            {
                throw new InvalidOperationException($"Parameter with index 0 is not defined");
            }
            var attributes = attributePath.Split('.');
            var attributeValue = Helper.GetValueFromAttributePath(entity, attributes, _orgContext);
            return attributeValue;
        }
    }

    public class FieldFunction : CrmBaseFunction
    {
        public FieldFunction(OrganizationServiceContext context)
            : base(context, "Field")
        {
        }

        

        public override object Evaluate(IExpression[] parameters, ExpressiveOptions options)
        {
            if (parameters.Length < 1)
            {
                throw new InvalidOperationException($"Function {Name} must have one parameter");
            }
            var entity = (Entity)Variables["entity"];
            var attributePath = parameters[0].Evaluate(Variables)?.ToString();
            if (string.IsNullOrEmpty(attributePath))
            {
                throw new InvalidOperationException($"Parameter with index 0 is not defined");
            }
            var attributes = attributePath.Split('.');
            var attributeValue = Helper.GetValueFromAttributePath(entity, attributes, _orgContext);
            if (attributeValue is EntityReference)
            {
                return ((EntityReference)attributeValue).Id;
            }
            else
            {
                return attributeValue;
            }
        }
    }

    public class CompareFunction : IFunction
    {
        public string Name { get { return "Compare"; } }

        public IDictionary<string, object> Variables { get; set; }

        public object Evaluate(IExpression[] parameters, ExpressiveOptions options)
        {
            if (parameters.Length < 2)
            {
                throw new InvalidOperationException($"Function {Name} must have two parameters");
            }
            var entity = (Entity)Variables["entity"];
            var preEntity = (Entity)Variables["preentity"];
            var part1 = parameters[0].Evaluate(Variables);
            var part2 = parameters[0].Evaluate(Variables);
            return (part1 == part2);
        }
    }

    public class FieldEntityFunction : CrmBaseFunction
    {
        public FieldEntityFunction(OrganizationServiceContext context)
            : base(context, "FieldEntity")
        {
        }

        public override object Evaluate(IExpression[] parameters, ExpressiveOptions options)
        {
            if (parameters.Length < 2)
            {
                throw new Exception($"Function {Name} must have two parameters");
            }
            var entityName = parameters[0].Evaluate(Variables).ToString();
            if (!Variables.ContainsKey(entityName))
            {
                throw new InvalidOperationException($"Entity with {entityName} name is not found in variables of expression");
            }
            var entity = Variables[entityName] as Entity;
            var attributePath = parameters[1].Evaluate(Variables).ToString();
            if (string.IsNullOrEmpty(attributePath))
            {
                throw new InvalidOperationException($"Parameter with index 0 is not defined");
            }
            var attributes = attributePath.Split('.');
            var attributeValue = Helper.GetValueFromAttributePath(entity, attributes, _orgContext);
            if (attributeValue is EntityReference)
            {
                return ((EntityReference)attributeValue).Id;
            }
            else
            {
                return attributeValue;
            }
        }
    }

    public class IsNullOrEmptyFunction : IFunction
    {
        public string Name { get { return "IsNullOrEmpty"; } }

        public IDictionary<string, object> Variables { get; set; }

        public object Evaluate(IExpression[] parameters, ExpressiveOptions options)
        {
            if (parameters.Length < 1)
            {
                throw new InvalidOperationException($"Function {Name} must have one parameter");
            }
            var @value = parameters[0].Evaluate(Variables) as string;
            return string.IsNullOrWhiteSpace(@value);
        }
    }

    public class JoinFunction : IFunction
    {
        public string Name { get { return "Join"; } }

        public IDictionary<string, object> Variables { get; set; }

        public object Evaluate(IExpression[] parameters, ExpressiveOptions options)
        {
            if (parameters.Length < 2)
            {
                throw new InvalidOperationException($"Function {Name} must have more than one parameter");
            }
            var separator = parameters[0].Evaluate(Variables).ToString();

            var result = new string[parameters.Length - 1];
            for(int i = 1;i < parameters.Length;i++)
            {
                result[i - 1] = parameters[i].Evaluate(Variables).ToString();
            }

            return string.Join(separator, result);
        }
    }

    internal class Helper
    {
        public static Expression Create(string template, params IFunction[] functions)
        {
            var parameterExpression = new Expression(template);
            if (functions != null)
            {
                foreach (var function in functions)
                {
                    parameterExpression.RegisterFunction(function);
                }
            }
            return parameterExpression;
        }


        public static object GetValueFromAttributePath(Entity baseEntity, string[] attributes, OrganizationServiceContext orgContext)
        {
            Entity entity = baseEntity;
            EntityReference previousAttribute = null;
            foreach (var attribute in attributes)
            {
                if (previousAttribute != null)
                {
                    orgContext.ClearChanges();
                    var pathAttributeValue = orgContext.CreateQuery(previousAttribute.LogicalName)
                        .Where(q => q.GetAttributeValue<Guid>($"{previousAttribute.LogicalName}id") == previousAttribute.Id)
                        .Select(q => q.Contains(attribute) ? q[attribute] : null)
                        .SingleOrDefault();

                    entity = new Entity(previousAttribute.LogicalName, previousAttribute.Id);
                    entity[attribute] = pathAttributeValue;
                }
                if (!entity.Contains(attribute))
                {
                    var baseAttributeValue = orgContext.CreateQuery(entity.LogicalName)
                        .Where(q => q.GetAttributeValue<Guid>($"{entity.LogicalName}id") == entity.Id)
                        .Select(q => q.Contains(attribute) ? q[attribute] : null)
                        .SingleOrDefault();
                    entity[attribute] = baseAttributeValue;
                }
                if (entity[attribute] == null)
                {
                    return null;
                }
                var attributeValue = entity[attribute];
                if (attributeValue is OptionSetValue)
                {
                    return ((OptionSetValue)attributeValue).Value;
                }
                else if (attributeValue is EntityReference)
                {
                    previousAttribute = (EntityReference)attributeValue;
                }
                else
                {
                    return attributeValue;
                }
            }
            return previousAttribute;
        }
    }
}
