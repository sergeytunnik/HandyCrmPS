﻿using System.Management.Automation;
using Microsoft.Xrm.Client;
using System;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsCommon.Get, "CRMConnection")]
	public class GetCrmConnectionCommand : PSCmdlet
	{
		[Parameter(
			Mandatory = true)]
		[ValidateNotNullOrEmpty]
		public string ConnectionString { get; set; }

		protected override void BeginProcessing()
		{
			base.BeginProcessing();

			CrmConnection crmConnection = CrmConnection.Parse(ConnectionString);
			crmConnection.ProxyTypesEnabled = false;
			// Defaul is PerName
			//crmConnection.ServiceConfigurationInstanceMode = ServiceConfigurationInstanceMode.PerName;

			WriteObject(crmConnection);
		}
	}
}