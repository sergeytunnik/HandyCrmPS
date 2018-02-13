using System.Management.Automation;
using System;
using Handy.Crm.Powershell.Cmdlets.Helpers;
using Handy.Crm.Powershell.Cmdlets.Connectivity;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsCommon.Get, "CRMConnection")]
    public class GetCrmConnectionCommand : PSCmdlet
    {
        [Parameter(
            Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public Uri OrganizationService { get; set; }

        [Parameter(Mandatory = true)]
        public PSCredential Credential { get; set; }

        [Parameter(
            Mandatory = false)]
        [ValidateNotNull]
        public Guid CallerId { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            var organizationService = new CrmService(
                new CrmServiceConfiguration(OrganizationService, Credential.ToClientCredentials()));

            if (MyInvocation.BoundParameters.ContainsKey("CallerId"))
                organizationService.CallerId = CallerId;

            WriteObject(organizationService);
        }
    }
}