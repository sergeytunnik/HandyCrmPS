using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Export, "CRMSolution")]
    public class ExportCrmSolutionCommand : CrmCmdletBase
    {
        private string AbsoluteFilePath
        {
            get
            {
                return System.IO.Path.IsPathRooted(FilePath)
                    ? FilePath
                    : System.IO.Path.GetFullPath(System.IO.Path.Combine(SessionState.Path.CurrentLocation.Path, FilePath));
            }
        }

        [Parameter(
            Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string FilePath { get; set; }

        [Parameter(
            Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string SolutionName { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter Managed { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportAutoNumberingSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportCalendarSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportCustomizationSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportEmailTrackingSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportGeneralSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportIsvConfig { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportMarketingSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportOutlookSynchronizationSettings { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportRelationshipRoles { get; set; }

        [Parameter(
            Mandatory = false)]
        public SwitchParameter ExportSales { get; set; }

        [Parameter(
            Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string TargetVersion { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            ExportSolutionRequest exportSolutionRequest = new ExportSolutionRequest()
            {
                Managed = Managed,
                SolutionName = SolutionName,
                ExportAutoNumberingSettings = ExportAutoNumberingSettings,
                ExportCalendarSettings = ExportCalendarSettings,
                ExportCustomizationSettings = ExportCustomizationSettings,
                ExportEmailTrackingSettings = ExportEmailTrackingSettings,
                ExportGeneralSettings = ExportGeneralSettings,
                ExportIsvConfig = ExportIsvConfig,
                ExportMarketingSettings = ExportMarketingSettings,
                ExportOutlookSynchronizationSettings = ExportOutlookSynchronizationSettings,
                ExportRelationshipRoles = ExportRelationshipRoles,
                ExportSales = ExportSales,
                TargetVersion = TargetVersion
            };

            WriteVerbose("Starting solution exporting");
            ExportSolutionResponse exportSolutionResponse = (ExportSolutionResponse)Connection.Execute(exportSolutionRequest);

            WriteVerbose(string.Format("Saving solution ({0}) at {1}", SolutionName, AbsoluteFilePath));
            File.WriteAllBytes(AbsoluteFilePath, exportSolutionResponse.ExportSolutionFile);
        }
    }
}
