using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	[Cmdlet(VerbsData.Export, "CRMSolution")]
	public class ExportCrmSolutionCommand : CrmCmdletBase
	{
		private string _directoryName;

		[Parameter(
			Mandatory = true)]
		[ValidateNotNullOrEmpty]
		public string SolutionName { get; set; }

		[Parameter(
			Mandatory = false)]
		public SwitchParameter Managed { get; set; }

		[Parameter(
			Mandatory = false)]
		[ValidateNotNullOrEmpty]
		public string DirectoryName
		{
			get
			{
				return _directoryName;
			}

			set
			{
				_directoryName = Path.IsPathRooted(value)
					? value
					: Path.GetFullPath(SessionState.Path.CurrentLocation.Path + value);
			}
		}

		[Parameter(
			Mandatory = false)]
		[ValidateNotNullOrEmpty]
		public string SolutionFileName { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest()
			{
				Managed = Managed,
				SolutionName = SolutionName,
			};

			WriteVerbose("Starting solution exporting");
			ExportSolutionResponse exportSolutionResponse = (ExportSolutionResponse)organizationService.Execute(exportSolutionRequest);

			var solutionXml = exportSolutionResponse.ExportSolutionFile;

			if (string.IsNullOrEmpty(DirectoryName))
				DirectoryName = SessionState.Path.CurrentLocation.Path;

			if (string.IsNullOrEmpty(SolutionFileName))
				SolutionFileName = string.Format("{0}{1}.zip", SolutionName, Managed ? "_managed" : "");

			string solutionPath = Path.Combine(DirectoryName, SolutionFileName);

			WriteVerbose(string.Format("Saving solution ({0}) at {1}", SolutionName, solutionPath));
			File.WriteAllBytes(solutionPath, solutionXml);
		}
	}
}
