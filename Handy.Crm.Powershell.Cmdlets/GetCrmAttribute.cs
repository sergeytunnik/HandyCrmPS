using System;
using System.Management.Automation;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsCommon.Get, "CRMAttribute")]
	public class GetCrmAttribute : CrmCmdletBase
	{
		[Parameter(
			Mandatory = true)]
		[ValidateNotNullOrEmpty]
		public string EntityName { get; set; }

		[Parameter(
			Mandatory = true)]
		[ValidateNotNullOrEmpty]
		public string AttributeName { get; set; }

		[Parameter(
			Mandatory = false,
			HelpMessage = "Set this value to true to include unpublished changes, as it would look if you called publish.\r\nSet this value to false to include only the currently published changes, ignoring the changes that haven't yet been published.")]
		public SwitchParameter RetrieveAsIfPublished { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest()
			{
				EntityLogicalName = EntityName,
				LogicalName = AttributeName,
				RetrieveAsIfPublished = RetrieveAsIfPublished
			};

			RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)organizationService.Execute(retrieveAttributeRequest);

			WriteObject(retrieveAttributeResponse);
		}
	}
}
