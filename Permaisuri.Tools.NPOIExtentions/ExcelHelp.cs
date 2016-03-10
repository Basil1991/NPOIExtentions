using System.Data;
using Permaisuri.Tools.NPOIExtentions.Argument;
using System.Collections.Generic;

namespace Permaisuri.Tools.NPOIExtentions {
    public class ExcelHelp {
        public void GetXls(DataTable dt, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(dt, arg);
        }
        public void GetXls(DataSet ds, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(ds, arg);
        }
        public void GetXls(dynamic list, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(list, arg);
        }
        public void GetXls(List<dynamic> lists, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(lists, arg);
        }
        public void GetXlsx(DataTable dt, ExcelArgument arg) {
            ExcelXlsxService.Export(dt, arg);
        }
        public void GetXlsx(DataSet ds, ExcelArgument arg) {
            ExcelXlsxService.Export(ds, arg);
        }
        public void GetXlsx(dynamic list, ExcelArgument arg) {
            ExcelXlsxService.Export(list, arg);
        }
        public void GetXlsx(List<dynamic> lists, ExcelArgument arg) {
            ExcelXlsxService.Export(lists, arg);
        }
    }
}
