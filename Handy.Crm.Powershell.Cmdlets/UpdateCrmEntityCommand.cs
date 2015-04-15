using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsData.Update, "CRMEntity")]
	public class UpdateCrmEntityCommand : CrmExecuteMultipleCmdletBase
	{
		[Parameter(
			Mandatory = true,
			ValueFromPipeline = true)]
		[ValidateNotNullOrEmpty]
		public Entity[] Entity { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			WriteVerbose("Entering in ProcessRecord");

			foreach (Entity e in Entity)
			{
				WriteVerbose(string.Format("Adding {0} ({1}) into ExecuteMultipleRequest", e.LogicalName, e.Id));

				e.UnwrapAttributes();
				UpdateRequest updateRequest = new UpdateRequest { Target = e };
				_executeMultipleRequest.Requests.Add(updateRequest);
			}
		}
	}
}