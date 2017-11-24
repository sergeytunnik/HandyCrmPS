using System;
using System.Collections;
using System.Management.Automation;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Handy.Crm.Powershell.Cmdlets.Helpers;

namespace Handy.Crm.Powershell.Cmdlets
{
  [Cmdlet(VerbsData.Update, "CRMEntity")]
  public class UpdateCrmEntityCommand : CrmExecuteMultipleCmdletBase
  {
    private const string ParameterSetNameEntity = "Entity";
    private const string ParameterSetNameHashtable = "Hashtable";

    [Parameter(
      Mandatory = true,
      ValueFromPipeline = true,
      ParameterSetName = ParameterSetNameEntity)]
    [ValidateNotNullOrEmpty]
    public Entity[] Entity { get; set; }

    [Parameter(
      Mandatory = true,
      ParameterSetName = ParameterSetNameHashtable)]
    [ValidateNotNullOrEmpty]
    public string EntityName { get; set; }

    [Parameter(
      Mandatory = true,
      ParameterSetName = ParameterSetNameHashtable)]
    [ValidateNotNull]
    public Guid Id { get; set; }

    [Parameter(
      Mandatory = true,
      ParameterSetName = ParameterSetNameHashtable)]
    [ValidateNotNull]
    public Hashtable Attributes { get; set; }

    protected override void ProcessRecord()
    {
      base.ProcessRecord();

      WriteVerbose("Entering in ProcessRecord");

      switch (ParameterSetName)
      {
        case ParameterSetNameEntity:
          foreach (var e in Entity)
          {
            WriteVerbose($"Adding {e.LogicalName} ({e.Id}) into ExecuteMultipleRequest");

            e.UnwrapAttributes();

            _executeMultipleRequest.Requests.Add(new UpdateRequest { Target = e });
          }
          break;

        case ParameterSetNameHashtable:
          WriteVerbose($"Adding {EntityName} ({Id}) into ExecuteMultipleRequest");

          var entityToUpdate = new Entity(EntityName, Id);

          foreach (string key in Attributes.Keys)
          {
            entityToUpdate[key] = Attributes[key].UnwrapPSObject();
          }

          _executeMultipleRequest.Requests.Add(new UpdateRequest {Target = entityToUpdate});
          break;
      }


    }
  }
}