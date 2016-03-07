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
    }
}
