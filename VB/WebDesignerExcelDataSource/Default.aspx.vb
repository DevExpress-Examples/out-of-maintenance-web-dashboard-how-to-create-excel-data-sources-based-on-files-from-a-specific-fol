Imports System
Imports System.Web.Hosting
Imports DevExpress.DashboardCommon
Imports DevExpress.DashboardWeb
Imports DevExpress.DataAccess.Excel
Imports DevExpress.Spreadsheet
Imports System.Linq
Imports System.IO

Namespace WebDesignerExcelDataSource
    Partial Public Class [Default]
        Inherits System.Web.UI.Page

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
            Dim dashboardFileStorage As New DashboardFileStorage("~/App_Data/Dashboards")
            ASPxDashboard1.SetDashboardStorage(dashboardFileStorage)

            Dim dataSourceStorage As New DataSourceInMemoryStorage()
            Using workbook As New Workbook()
                Directory.EnumerateFiles(HostingEnvironment.MapPath("~/App_Data/ExcelFiles/"), "*.xlsx").SelectMany(Function(file)
                    workbook.LoadDocument(file)
                    Return workbook.Worksheets.Select(Function(sheet)
                        Dim dataSourceName = String.Format("{0} - {1}", Path.GetFileNameWithoutExtension(file), sheet.Name)
                        Dim excelDataSource = New DashboardExcelDataSource(dataSourceName)
                        excelDataSource.FileName = file
                        Dim worksheetSettings = New ExcelWorksheetSettings() With {.WorksheetName = sheet.Name}
                        excelDataSource.SourceOptions = New ExcelSourceOptions(worksheetSettings)
                        Return New With { _
                            Key .Name = excelDataSource.Name, _
                            Key .Xml = excelDataSource.SaveToXml() _
                        }
                    End Function)
                End Function).ToList().ForEach(Sub(ds) dataSourceStorage.RegisterDataSource(ds.Name, ds.Xml))
            End Using
            ASPxDashboard1.SetDataSourceStorage(dataSourceStorage)
        End Sub
    End Class
End Namespace