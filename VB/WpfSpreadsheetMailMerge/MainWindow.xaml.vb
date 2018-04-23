Imports DevExpress.Spreadsheet
Imports DevExpress.Xpf.Core
Imports DevExpress.Xpf.Spreadsheet
Imports System.Collections.Generic
Imports System

Namespace WpfSpreadsheetMailMerge
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits DevExpress.Xpf.Core.ThemedWindow

        Private previewWindow As ThemedWindow
        Public Sub New()
            InitializeComponent()
            CreateTemplate()
        End Sub

        Private ReadOnly Property Workbook() As IWorkbook
            Get
                Return spreadsheetControl.Document
            End Get
        End Property
        Private ReadOnly Property MailMergeMode() As DefinedName
            Get
                Return Workbook.DefinedNames.GetDefinedName("MAILMERGEMODE")
            End Get
        End Property

        Public Sub CreateTemplate()
            ' Create a mail merge template.
            Dim template As Worksheet = Workbook.Worksheets(0)
            template.Cells("C2").Formula = "FIELDPICTURE(""Photo"", ""range"", C2, FALSE, 50)"
            template.Cells("C3").Formula = "=FIELD(""FirstName"")&"" ""&FIELD(""LastName"")"
            template.Cells("B4").Value = "Position:"
            template.Cells("C4").Formula = "FIELD(""Title"")"
            template.Cells("B5").Value = "Birth Date:"
            template.Cells("C5").Formula = "FIELD(""BirthDate"")"
            template.Cells("C5").NumberFormat = "MMMM dd, yyyy"
            template.Cells("B6").Value = "Hire Date:"
            template.Cells("C6").Formula = "FIELD(""HireDate"")"
            template.Cells("C6").NumberFormat = "MMMM dd, yyyy"
            template.Cells("B7").Value = "Home Phone:"
            template.Cells("C7").Formula = "FIELD(""HomePhone"")"
            template.Cells("B8").Value = "Address:"
            template.Cells("C8").Formula = "=FIELD(""Address"")&"" ""&FIELD(""City"")"
            template.Cells("B9").Value = "About:"
            template.Cells("C9").Formula = "FIELD(""Notes"")"

            ' Set a detail range in the template.
            Dim detail As Range = template.Range("C1:C9")
            detail.Name = "DETAILRANGE"

            ' Set a header range in the template.
            Dim header As Range = template.Range("B1:B9")
            header.Name = "HEADERRANGE"

            SetDefaultMailMergeOptions()
        End Sub

        Public Sub SetDefaultMailMergeOptions()
            ' Set the mail merge mode to "Multiple Sheets".
            Workbook.DefinedNames.Add("MAILMERGEMODE", """Worksheets""")

            ' Set the document orientation.
            Workbook.DefinedNames.Add("HORIZONTALMODE", "TRUE")
        End Sub

        Public Sub ShowPreview()
            If previewWindow IsNot Nothing Then
                previewWindow.Close()
            End If

            PerformMailMerge()
            previewWindow = New ThemedWindow()
            previewWindow.Owner = Me
            previewWindow.Title = "Mail Merge Preview"
            Dim control As New SpreadsheetControl()
            control.Document.LoadDocument("MailMergeResult.xlsx")
            previewWindow.Content = control
            previewWindow.Show()
        End Sub

        Public Sub PerformMailMerge()
            ' Perform a mail merge.
            Workbook.MailMergeDataSource = EmployeesInfo.GetData()
            Dim result As IList(Of IWorkbook) = Workbook.GenerateMailMergeDocuments()
            result(0).SaveDocument("MailMergeResult.xlsx")
        End Sub

        Public Function CanSelectMode() As Boolean
            Return MailMergeMode IsNot Nothing
        End Function

        Public Sub SelectSingleSheetMode()
            MailMergeMode.RefersTo = """OneWorksheet"""
        End Sub

        Public Sub SelectMultipleSheetsMode()
            MailMergeMode.RefersTo = """Worksheets"""
        End Sub

        Public Sub SelectMultipleDocumentsMode()
            MailMergeMode.RefersTo = """Documents"""
        End Sub
    End Class
End Namespace
