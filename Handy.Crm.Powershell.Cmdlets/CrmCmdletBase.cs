using System;
using System.Management.Automation;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
	public class CrmCmdletBase : PSCmdlet
	{
		protected IOrganizationService organizationService { get; set; }

		[Parameter(
			Mandatory = true,
			Position = 0)]
		[ValidateNotNull]
		public CrmServiceClient Connection { get; set; }

		protected override void BeginProcessing()
		{
			base.BeginProcessing();

            organizationService = Connection;
		}

		protected override void EndProcessing()
		{
			base.EndProcessing();
		}
	}
}
