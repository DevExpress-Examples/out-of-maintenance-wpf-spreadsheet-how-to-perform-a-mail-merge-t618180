using DevExpress.Spreadsheet;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Spreadsheet;
using System.Collections.Generic;
using System;

namespace WpfSpreadsheetMailMerge
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DevExpress.Xpf.Core.ThemedWindow
    {
        ThemedWindow previewWindow;
        public MainWindow()
        {
            InitializeComponent();
            CreateTemplate();
        }

        IWorkbook Workbook { get { return spreadsheetControl.Document; } }
        DefinedName MailMergeMode { get { return Workbook.DefinedNames.GetDefinedName("MAILMERGEMODE"); } }
        
        public void CreateTemplate()
        {
            // Create a mail merge template.
            Worksheet template = Workbook.Worksheets[0];
            template.Cells["C2"].Formula = "FIELDPICTURE(\"Photo\", \"range\", C2, FALSE, 50)";
            template.Cells["C3"].Formula = "=FIELD(\"FirstName\")&\" \"&FIELD(\"LastName\")";
            template.Cells["B4"].Value = "Position:";
            template.Cells["C4"].Formula = "FIELD(\"Title\")";
            template.Cells["B5"].Value = "Birth Date:";
            template.Cells["C5"].Formula = "FIELD(\"BirthDate\")";
            template.Cells["C5"].NumberFormat = "MMMM dd, yyyy";
            template.Cells["B6"].Value = "Hire Date:";
            template.Cells["C6"].Formula = "FIELD(\"HireDate\")";
            template.Cells["C6"].NumberFormat = "MMMM dd, yyyy";
            template.Cells["B7"].Value = "Home Phone:";
            template.Cells["C7"].Formula = "FIELD(\"HomePhone\")";
            template.Cells["B8"].Value = "Address:";
            template.Cells["C8"].Formula = "=FIELD(\"Address\")&\" \"&FIELD(\"City\")";
            template.Cells["B9"].Value = "About:";
            template.Cells["C9"].Formula = "FIELD(\"Notes\")";

            // Set a detail range in the template.
            Range detail = template.Range["C1:C9"];
            detail.Name = "DETAILRANGE";

            // Set a header range in the template.
            Range header = template.Range["B1:B9"];
            header.Name = "HEADERRANGE";

            SetDefaultMailMergeOptions();
        }

        public void SetDefaultMailMergeOptions()
        {
            // Set the mail merge mode to "Multiple Sheets".
            Workbook.DefinedNames.Add("MAILMERGEMODE", "\"Worksheets\"");

            // Set the document orientation.
            Workbook.DefinedNames.Add("HORIZONTALMODE", "TRUE");
        }

        public void ShowPreview()
        {
            if (previewWindow != null)
                previewWindow.Close();

            PerformMailMerge();
            previewWindow = new ThemedWindow();
            previewWindow.Owner = this;
            previewWindow.Title = "Mail Merge Preview";
            SpreadsheetControl control = new SpreadsheetControl();
            control.Document.LoadDocument("MailMergeResult.xlsx");
            previewWindow.Content = control;
            previewWindow.Show();
        }

        public void PerformMailMerge()
        {
            // Perform a mail merge.
            Workbook.MailMergeDataSource = EmployeesInfo.GetData();
            IList<IWorkbook> result = Workbook.GenerateMailMergeDocuments();
            result[0].SaveDocument("MailMergeResult.xlsx");
        }

        public bool CanSelectMode()
        {
            return MailMergeMode != null;
        }

        public void SelectSingleSheetMode()
        {
            MailMergeMode.RefersTo = "\"OneWorksheet\"";
        }

        public void SelectMultipleSheetsMode()
        {
            MailMergeMode.RefersTo = "\"Worksheets\"";
        }

        public void SelectMultipleDocumentsMode()
        {
            MailMergeMode.RefersTo = "\"Documents\"";
        }
    }
}
