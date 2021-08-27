<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128612839/17.1.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T618180)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [EmployeeInfo.cs](./CS/WpfSpreadsheetMailMerge/EmployeeInfo.cs) (VB: [EmployeeInfo.vb](./VB/WpfSpreadsheetMailMerge/EmployeeInfo.vb))
* **[MainWindow.xaml](./CS/WpfSpreadsheetMailMerge/MainWindow.xaml) (VB: [MainWindow.xaml](./VB/WpfSpreadsheetMailMerge/MainWindow.xaml))**
* [MainWindow.xaml.cs](./CS/WpfSpreadsheetMailMerge/MainWindow.xaml.cs) (VB: [MainWindow.xaml.vb](./VB/WpfSpreadsheetMailMerge/MainWindow.xaml.vb))
<!-- default file list end -->
# WPF Spreadsheet - How to perform a mail merge


This example demonstrates how to use the <strong>Spreadsheet Mail Merge</strong> functionality to automatically generate a document based on data retrieved from a data source. <br>The mail merge template is created in code when the application starts. The <a href="https://documentation.devexpress.com/CoreLibraries/DevExpress.Spreadsheet.IWorkbook.MailMergeDataSource.property">IWorkbook.MailMergeDataSource</a> property specifies the data source. The <a href="https://documentation.devexpress.com/CoreLibraries/DevExpress.Spreadsheet.IWorkbook.GenerateMailMergeDocuments.method">IWorkbook.GenerateMailMergeDocuments</a> method accomplishes mail merge and returns the resulting workbook.<br>The <strong>Mail Merge</strong> tab is added to the SpreadsheetControl's ribbon to provide end-users with capabilities to select the mail merge mode and preview the merged document. If the <strong>Multiple Documents</strong> mode is used, only the first merged workbook is shown in the preview window.<br><img src="https://raw.githubusercontent.com/DevExpress-Examples/wpf-spreadsheet-how-to-perform-a-mail-merge-t618180/17.1.3+/media/c9160390-a446-4e31-b83a-b290731e5fa6.png">

<br/>


