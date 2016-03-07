using System.Data;
using Permaisuri.Tools.NPOIExtentions.Argument;
using System.Collections.Generic;

namespace Permaisuri.Tools.NPOIExtentions {
    public class ExcelHelp {
        public void Get(DataTable dt, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(dt, arg);
        }
        public void Get(DataSet ds, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(ds, arg);
        }
        public void Get(dynamic list, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(list, arg);
        }
        public void Get(List<dynamic> lists, ExcelArgument arg) {
            ExcelXlsService.ExportXlsExcel(lists, arg);
        }
    }
}
