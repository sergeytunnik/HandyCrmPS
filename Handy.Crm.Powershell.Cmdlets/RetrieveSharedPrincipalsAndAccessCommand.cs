using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsCommon.Get, "CRMSharedPrincipalsAndAccess")]
	public class RetrieveSharedPrincipalsAndAccessCommand : CrmCmdletBase
	{
		[Parameter(Mandatory = true)]
		public EntityReference Target { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			RetrieveSharedPrincipalsAndAccessRequest request = new RetrieveSharedPrincipalsAndAccessRequest()
			{
				Target = Target
			};

			WriteVerbose("Executing RetrieveSharedPrincipalsAndAccessRequest");
			RetrieveSharedPrincipalsAndAccessResponse response = (RetrieveSharedPrincipalsAndAccessResponse)organizationService.Execute(request);

			WriteObject(response);
		}
	}
}
