using System.Management.Automation;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
    public abstract class CrmCmdletBase : PSCmdlet
    {
        [Parameter(
            Mandatory = true,
            Position = 0)]
        [ValidateNotNull]
        public IOrganizationService Connection { get; set; }
    }
}
