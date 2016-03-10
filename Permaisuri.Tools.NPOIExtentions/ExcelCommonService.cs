using NPOI.SS.UserModel;
using System.IO;

namespace Permaisuri.Tools.NPOIExtentions {
    public class ExcelCommonService {
        public static void WriteToFile(IWorkbook hssfworkbook, string outPutPath) {
            //Write the stream data of workbook to the root directory
            using (FileStream file = new FileStream(outPutPath, FileMode.Create)) {
                hssfworkbook.Write(file);
            }
        }
    }
}
