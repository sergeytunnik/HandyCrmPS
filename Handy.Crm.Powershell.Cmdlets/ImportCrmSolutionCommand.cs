﻿using System;
using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsData.Import, "CRMSolution")]
	public class ImportCrmSolutionCommand : CrmCmdletBase
	{
		private string _path;

		[Parameter(
			Mandatory = true),
		ValidateNotNullOrEmpty]
		public string Path
		{
			get
			{
				return _path;
			}

			set
			{
				_path = System.IO.Path.IsPathRooted(value)
					? value
					: System.IO.Path.GetFullPath(SessionState.Path.CurrentLocation.Path + value);
			}
		}

		[Parameter(
			Mandatory = false)]
		public SwitchParameter OverwriteCustomizations { get; set; }

		[Parameter(
			Mandatory = false)]
		public SwitchParameter PublishWorkflows { get; set; }

		[Parameter(
			Mandatory = false)]
		public SwitchParameter PassThru { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			byte[] fileBytes = File.ReadAllBytes(Path);

			ImportSolutionRequest impSolutionRequest = new ImportSolutionRequest()
			{
				CustomizationFile = fileBytes,
				ImportJobId = Guid.NewGuid(),
				OverwriteUnmanagedCustomizations = OverwriteCustomizations,
				PublishWorkflows = PublishWorkflows
			};

			WriteVerbose(string.Format("Start importing solution from {0}", Path));
			organizationService.Execute(impSolutionRequest);
			WriteVerbose("Finished importing solution");

			if (PassThru)
				WriteObject(impSolutionRequest.ImportJobId);
		}
	}
}
