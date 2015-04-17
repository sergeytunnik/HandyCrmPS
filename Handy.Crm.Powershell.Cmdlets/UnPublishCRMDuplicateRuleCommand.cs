using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsData.Unpublish, "CRMDuplicateRule")]
	public class UnpublishCRMDuplicateRuleCommand : CrmCmdletBase
	{
		[Parameter(Mandatory = true)]
		public Guid Id { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			UnpublishDuplicateRuleRequest unPublishDuplicateRuleRequest = new UnpublishDuplicateRuleRequest()
			{
				DuplicateRuleId = Id
			};

			WriteVerbose("Executing PublishDuplicateRuleRequest");
			UnpublishDuplicateRuleResponse response = (UnpublishDuplicateRuleResponse)organizationService.Execute(unPublishDuplicateRuleRequest);

			WriteObject(response);
		}
	}
}
