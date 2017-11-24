using System.Management.Automation;
using System;
using Microsoft.Xrm.Sdk.Client;
using Handy.Crm.Powershell.Cmdlets.Helpers;

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

            var organizationService = new OrganizationServiceProxy(OrganizationService, null, Credential.ToClientCredentials(), null);

            if (MyInvocation.BoundParameters.ContainsKey("CallerId"))
                organizationService.CallerId = CallerId;

            WriteObject(organizationService);
        }
    }
}