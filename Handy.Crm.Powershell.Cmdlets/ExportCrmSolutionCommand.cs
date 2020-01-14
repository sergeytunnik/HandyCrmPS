using System.IO;
using System.Management.Automation;
using Microsoft.Crm.Sdk.Messages;

namespace Handy.Crm.Powershell.Cmdlets
{
    [Cmdlet(VerbsData.Export, "CRMSolution")]
    public class ExportCrmSolutionCommand : CrmCmdletBase
    {
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
            var exportSolutionRequest = new ExportSolutionRequest()
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
            var exportSolutionResponse = (ExportSolutionResponse)Connection.Execute(exportSolutionRequest);

            var absoluteFilePath = GetAbsoluteFilePath(FilePath);
            WriteVerbose(string.Format("Saving solution ({0}) at {1}", SolutionName, absoluteFilePath));
            File.WriteAllBytes(absoluteFilePath, exportSolutionResponse.ExportSolutionFile);
        }
    }
}
