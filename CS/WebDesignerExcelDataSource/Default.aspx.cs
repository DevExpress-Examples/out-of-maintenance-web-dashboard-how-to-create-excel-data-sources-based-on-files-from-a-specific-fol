using System;
using System.Web.Hosting;
using DevExpress.DashboardCommon;
using DevExpress.DashboardWeb;
using DevExpress.DataAccess.Excel;
using DevExpress.Spreadsheet;
using System.Linq;
using System.IO;

namespace WebDesignerExcelDataSource {
    public partial class Default : System.Web.UI.Page {
        protected void Page_Load(object sender, EventArgs e) {
            DashboardFileStorage dashboardFileStorage = new DashboardFileStorage("~/App_Data/Dashboards");
            ASPxDashboard1.SetDashboardStorage(dashboardFileStorage);

            DataSourceInMemoryStorage dataSourceStorage = new DataSourceInMemoryStorage();
            using (Workbook workbook = new Workbook()) {
                Directory
                .EnumerateFiles(HostingEnvironment.MapPath(@"~/App_Data/ExcelFiles/"), "*.xlsx")
                .SelectMany(file => {
                    workbook.LoadDocument(file);
                    return workbook.Worksheets.Select(sheet => {
                        var dataSourceName = string.Format("{0} - {1}", Path.GetFileNameWithoutExtension(file), sheet.Name);
                        var excelDataSource = new DashboardExcelDataSource(dataSourceName);
                        excelDataSource.FileName = file;
                        var worksheetSettings = new ExcelWorksheetSettings() { WorksheetName = sheet.Name };
                        excelDataSource.SourceOptions = new ExcelSourceOptions(worksheetSettings);
                        return new {
                            Name = excelDataSource.Name,
                            Xml = excelDataSource.SaveToXml()
                        };
                    });
                })
               .ToList()
               .ForEach(ds => dataSourceStorage.RegisterDataSource(ds.Name, ds.Xml));
            }
            ASPxDashboard1.SetDataSourceStorage(dataSourceStorage);
        }
    }
}