using System.Management.Automation;
using Microsoft.Xrm.Tooling.Connector;
using System;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsCommon.Get, "CRMConnection")]
    public class GetCrmConnectionCommand : PSCmdlet
    {
        [Parameter(
            Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string ConnectionString { get; set; }

        [Parameter(
            Mandatory = false)]
        [ValidateNotNull]
        public Guid CallerId { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            var crmServiceClient = new CrmServiceClient(ConnectionString);

            if (MyInvocation.BoundParameters.ContainsKey("CallerId"))
                crmServiceClient.CallerId = CallerId;

            WriteObject(crmServiceClient);
        }
    }
}