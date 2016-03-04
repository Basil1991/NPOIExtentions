using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Permaisuri.Tools.NPOIExtentions.Argument;

namespace Permaisuri.Tools.NPOIExtentions {
    internal class StyleProcesser {
        private static ICellStyle defaultStyle = null;
        public static ICellStyle GetStyleByClass(ClassType type, HSSFWorkbook hssfworkbook) {
            switch (type) {
                case ClassType.Default:
                    if (defaultStyle == null) {
                        ICellStyle style = hssfworkbook.CreateCellStyle();
                        style.WrapText = true;

                        //style.BorderBottom = BorderStyle.Thin;
                        //style.BottomBorderColor = HSSFColor.Black.Index;
                        //style.BorderLeft = BorderStyle.DashDotDot;
                        //style.LeftBorderColor = HSSFColor.Green.Index;
                        //style.BorderRight = BorderStyle.Hair;
                        //style.RightBorderColor = HSSFColor.Blue.Index;
                        //style.BorderTop = BorderStyle.MediumDashed;
                        //style.TopBorderColor = HSSFColor.Orange.Index;
                        defaultStyle = style;
                    }
                    return defaultStyle;
                default: return null;
            }
        }

        private static ICellStyle datetimeStyle = null;
        public static ICellStyle GetCellStyleByDataType(ColumnValueType type, HSSFWorkbook hssfworkbook) {
            switch (type) {
                case ColumnValueType.DateTime:
                    if (datetimeStyle == null) {
                        ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
                        IDataFormat format = hssfworkbook.CreateDataFormat();
                        cellStyle.DataFormat = format.GetFormat("yyyy/m/d  h:mm:sss");

                        datetimeStyle = cellStyle;
                    }
                    return datetimeStyle;
                //case ColumnValueType.Picture:
                //    if (pictureStyle == null) {
                //        ICellStyle cellStyle = hssfworkbook.CreateCellStyle();
                //    }
                default: break;
            }
            return null;
        }
    }
}
