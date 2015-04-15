using System;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsCommon.Remove, "CRMEntity",
		SupportsShouldProcess = true)]
	public class RemoveCrmEntityCommand : CrmExecuteMultipleCmdletBase
	{
		[Parameter(
			Mandatory = true,
			ParameterSetName = "EntityNameAndId")]
		[ValidateNotNullOrEmpty]
		public string EntityName { get; set; }

		[Parameter(
			Mandatory = true,
			ParameterSetName = "EntityNameAndId",
			ValueFromPipeline = true)]
		public Guid[] Id { get; set; }

		[Parameter(
			Mandatory = true,
			ParameterSetName = "Entity",
			ValueFromPipeline = true)]
		[ValidateNotNullOrEmpty]
		public Entity[] Entity { get; set; }


		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			WriteVerbose("Entering in ProcessRecord");

			switch (ParameterSetName)
			{
				case "Entity":
					WriteVerbose("Entering Entity parameter set");

					foreach (Entity e in Entity)
					{
						WriteVerbose(string.Format("Adding {0} ({1}) into ExecuteMultipleRequest", e.LogicalName, e.Id));

						AddDeleteRequestWithConfirmation(e.LogicalName, e.Id);
					}
					break;

				case "EntityNameAndId":
					WriteVerbose("Entering EntityNameAndId parameter set");

					foreach (Guid g in Id)
					{
						WriteVerbose(string.Format("Adding {0} ({1}) into ExecuteMultipleRequest", EntityName, g));

						AddDeleteRequestWithConfirmation(EntityName, g);
					}
					break;

				default:
					WriteVerbose("Entering Default parameter set");
					break;
			}
		}

		private void AddDeleteRequestWithConfirmation(string entityName, Guid id)
		{
			
			if (ShouldProcess(string.Format("{0} ({1})", entityName, id)))
			{
				_executeMultipleRequest.Requests.Add(
					new DeleteRequest
					{
						Target = new EntityReference(entityName, id)
					});
			}
			else
			{
				WriteVerbose(string.Format("Skipping {0} ({1})", entityName, id));
			}
		}
	}
}