using NPOI.SS.UserModel;
using Permaisuri.Tools.NPOIExtentions.Argument;

namespace Permaisuri.Tools.NPOIExtentions {
    internal class StyleProcesser {
        public static ICellStyle GetStyleByClass(ClassType type, IWorkbook workbook) {
            switch (type) {
                case ClassType.Default:
                    ICellStyle style = workbook.CreateCellStyle();
                    style.WrapText = true;
                    style.VerticalAlignment = VerticalAlignment.Center;
                    style.Alignment = HorizontalAlignment.Center;
                    return style;
                default: return null;
            }
        }
        public static void SetSheetDefaultStyle(ColumnValueType type, IWorkbook workbook, ISheet sheet, int columnIndex) {
            ICellStyle cellStyle;
            switch (type) {
                case ColumnValueType.DateTime:
                    cellStyle = workbook.CreateCellStyle();
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    IDataFormat format = workbook.CreateDataFormat();
                    cellStyle.DataFormat = format.GetFormat("yyyy/m/d  h:mm:sss");
                    sheet.SetDefaultColumnStyle(columnIndex, cellStyle);
                    break;
                default:
                    cellStyle = workbook.CreateCellStyle();
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    sheet.SetDefaultColumnStyle(columnIndex, cellStyle);
                    break;
            }
        }
    }
}
