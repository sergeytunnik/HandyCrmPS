using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
	public enum ExportTypeEnum
	{
		Managed, Unmanaged, Both
	};

	[Cmdlet(VerbsData.Export, "CRMSolution")]
	public class ExportCrmSolutionCommand : CrmCmdletBase
	{
		[Parameter(
			Mandatory = true)]
		public string Name { get; set; }

		[Parameter(
			Mandatory = true)]
		public ExportTypeEnum Type { get; set; }

		[Parameter(
			Mandatory = true)]
		public string DirectoryName { get; set; }

		protected override void ProcessRecord()
		{
			base.ProcessRecord();

			switch (Type)
			{
				case ExportTypeEnum.Managed:
					ExportSolutionHelper(Name, true);
					break;

				case ExportTypeEnum.Unmanaged:
					ExportSolutionHelper(Name, false);
					break;

				case ExportTypeEnum.Both:
					ExportSolutionHelper(Name, false);
					ExportSolutionHelper(Name, true);
					break;
			}
		}

		private void ExportSolutionHelper(string solutionName, bool isManaged)
		{
			ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest()
			{
				Managed = isManaged,
				SolutionName = solutionName,
			};

			WriteVerbose("Starting solution exporting");
			ExportSolutionResponse exportSolutionResponse = (ExportSolutionResponse)organizationService.Execute(exportSolutionRequest);

			var solutionXml = exportSolutionResponse.ExportSolutionFile;

			string solutionFileName = string.Format("{0}{1}.zip", solutionName, isManaged ? "_managed" : "");
			string solutionPath = Path.Combine(DirectoryName, solutionFileName);

			WriteVerbose(string.Format("Saving solution ({0}) at {1}", solutionName, solutionPath));
			File.WriteAllBytes(solutionPath, solutionXml);
		}
	}
}
