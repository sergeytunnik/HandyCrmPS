using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsData.Publish, "CRMDuplicateRule")]
	public class PublishCrmDuplicateRuleCommand : CrmCmdletBase
	{
		[Parameter(Mandatory = true)]
		public Guid Id { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			PublishDuplicateRuleRequest publishDuplicateRuleRequest = new PublishDuplicateRuleRequest()
			{
				DuplicateRuleId = Id
			};

			WriteVerbose("Executing PublishDuplicateRuleRequest");
			PublishDuplicateRuleResponse response = (PublishDuplicateRuleResponse)organizationService.Execute(publishDuplicateRuleRequest);

			WriteObject(response);
		}
	}
}
