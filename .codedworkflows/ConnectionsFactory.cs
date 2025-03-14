using UiPath.CodedWorkflows;
using System;

namespace RPA_AutoUpdateAnchor
{
    public class GoogleDocsFactory
    {
        public UiPath.GSuite.Activities.Api.DocsConnection My_Workspace_nguyendinhhai196_gmail_com__2 { get; set; }

        public UiPath.GSuite.Activities.Api.DocsConnection My_Workspace_haiiicoo1996_gmail_com__2 { get; set; }

        public GoogleDocsFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_nguyendinhhai196_gmail_com__2 = new UiPath.GSuite.Activities.Api.DocsConnection("898f3a88-8a22-4d3c-bb3a-ad206f04118f", resolver);
            My_Workspace_haiiicoo1996_gmail_com__2 = new UiPath.GSuite.Activities.Api.DocsConnection("91e90991-afb5-4f0b-9fad-dd894e8a33d6", resolver);
        }
    }

    public class DriveFactory
    {
        public UiPath.GSuite.Activities.Api.DriveConnection My_Workspace_nguyendinhhai196_gmail_com__2 { get; set; }

        public UiPath.GSuite.Activities.Api.DriveConnection My_Workspace_haiiicoo1996_gmail_com__2 { get; set; }

        public UiPath.GSuite.Activities.Api.DriveConnection My_Workspace_haiiicoo1996_gmail_com__3 { get; set; }

        public DriveFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_nguyendinhhai196_gmail_com__2 = new UiPath.GSuite.Activities.Api.DriveConnection("9ec480b7-fae9-4cdd-92ec-3d02a4f03c90", resolver);
            My_Workspace_haiiicoo1996_gmail_com__2 = new UiPath.GSuite.Activities.Api.DriveConnection("f78525f8-d247-46ca-a076-e91225396400", resolver);
            My_Workspace_haiiicoo1996_gmail_com__3 = new UiPath.GSuite.Activities.Api.DriveConnection("a1574d43-ee1e-49c8-9a7e-78b789cbc648", resolver);
        }
    }

    public class GmailFactory
    {
        public UiPath.GSuite.Activities.Api.GmailConnection My_Workspace_nguyendinhhai196_gmail_com__2 { get; set; }

        public GmailFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_nguyendinhhai196_gmail_com__2 = new UiPath.GSuite.Activities.Api.GmailConnection("f654b2e7-6ab0-4c96-844d-4bdd9c768158", resolver);
        }
    }

    public class GoogleSheetsFactory
    {
        public UiPath.GSuite.Activities.Api.SheetsConnection My_Workspace_nguyendinhhai196_gmail_com__2 { get; set; }

        public UiPath.GSuite.Activities.Api.SheetsConnection My_Workspace_haiiicoo1996_gmail_com__2 { get; set; }

        public GoogleSheetsFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_nguyendinhhai196_gmail_com__2 = new UiPath.GSuite.Activities.Api.SheetsConnection("f215dd52-8fe8-429f-bd09-bbaeb01166f9", resolver);
            My_Workspace_haiiicoo1996_gmail_com__2 = new UiPath.GSuite.Activities.Api.SheetsConnection("b15fdeef-4718-4039-88ab-00ec795e9084", resolver);
        }
    }
}