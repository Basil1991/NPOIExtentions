using System.Data;
using Permaisuri.Tools.NPOIExtentions.Argument;

namespace Permaisuri.Tools.NPOIExtentions {
    public class ExcelHelp {
        public void Get(DataTable dt, ExcelArgument arg) {
            ExcelService.ExportExcel(dt, arg);
        }
    }
}
