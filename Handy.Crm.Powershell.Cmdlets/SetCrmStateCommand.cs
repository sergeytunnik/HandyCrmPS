using System;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsCommon.Set, "CRMState")]
    public class SetCrmStateCommand : CrmExecuteMultipleCmdletBase
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

        [Parameter(
            Mandatory = true)]
        public OptionSetValue State { get; set; }

        [Parameter(
            Mandatory = true)]
        public OptionSetValue Status { get; set; }

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
                        AddSetStateRequest(e.LogicalName, e.Id, State, Status);
                    }
                    break;

                case "EntityNameAndId":
                    WriteVerbose("Entering EntityNameAndId parameter set");

                    foreach (Guid g in Id)
                    {
                        AddSetStateRequest(EntityName, g, State, Status);
                    }
                    break;
            }
        }

        private void AddSetStateRequest(string entityName, Guid id, OptionSetValue state, OptionSetValue status)
        {
            WriteVerbose(string.Format("Adding {0} ({1}) into ExecuteMultipleRequest", entityName, id));

            _executeMultipleRequest.Requests.Add(new SetStateRequest
            {
                EntityMoniker = new EntityReference(entityName, id),
                State = state,
                Status = status
            });
        }
    }
}
