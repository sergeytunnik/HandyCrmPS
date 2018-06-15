using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsSecurity.Revoke, "CRMAccess")]
    public class RevokeAccessCommand : CrmCmdletBase
    {
        [Parameter(Mandatory = true)]
        public EntityReference Revokee { get; set; }

        [Parameter(Mandatory = true)]
        public EntityReference Target { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            RevokeAccessRequest request = new RevokeAccessRequest()
            {
                Revokee = Revokee,
                Target = Target
            };

            WriteVerbose("Executing RevokeAccessResponse");
            RevokeAccessResponse response = (RevokeAccessResponse)Connection.Execute(request);

            WriteObject(response);
        }
    }
}
