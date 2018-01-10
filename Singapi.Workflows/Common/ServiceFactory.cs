using Microsoft.Xrm.Sdk;
using System.Text;

namespace Singapi.Workflows.Common
{
    class ServiceFactory
    {
        public StringBuilder Log { get; set; }

        public IOrganizationService Service { get; set; }
    }
}