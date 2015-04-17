﻿using System.Management.Automation;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;

namespace Handy.Crm.Powershell.Cmdlets
{
  public class CrmCmdletBase : PSCmdlet
  {
    protected OrganizationService organizationService { get; set; }

    [Parameter(
      Mandatory = true)]
    [ValidateNotNull]
    public CrmConnection Connection { get; set; }

    protected override void BeginProcessing()
    {
      base.BeginProcessing();

      WriteVerbose("Creating OrganizationService");
			organizationService = new OrganizationService(Connection);
    }

    protected override void EndProcessing()
    {
      WriteVerbose("Disposing OrganizationService");
      organizationService.Dispose();

			base.EndProcessing();
    }
  }
}
