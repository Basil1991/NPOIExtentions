using System;
using System.Collections.Generic;

namespace Permaisuri.Tools.NPOIExtentions.Argument {
    public class ExcelArgument {
        public ExcelArgument(string outPutPath, List<SheetArgument> sheetArguments) {
            if (string.IsNullOrWhiteSpace(outPutPath))
                throw new ArgumentException("outPutPath");
            this.OutPutPath = outPutPath;
            if (sheetArguments == null || sheetArguments.Count <= 0) {
                throw new ArgumentException("sheetArguments");
            }
            this.SheetArguments = sheetArguments;
        }
        public string OutPutPath { get; }
        public List<SheetArgument> SheetArguments { get; }
    }
    public class SheetArgument {
        public SheetArgument(ColumnArgument[] columnArguments, string sheetName, short height = 2, bool isTitleShow = false, ClassType classType = ClassType.Default) {
            if (columnArguments == null) {
                throw new ArgumentException("columnArguments");
            }
            this.ColumnArguments = columnArguments;
            if (string.IsNullOrWhiteSpace(sheetName)) {
                this.SheetName = Guid.NewGuid().ToString();
            }
            else {
                this.SheetName = sheetName;
            }
            this.ClassType = classType;
            this.TitleHeight = 170 * 3;
            this.RowHeight = Convert.ToInt16(170 * height);
        }
        public string ColumnHeight { get; }
        public short TitleHeight { get; }
        public short RowHeight { get; }
        public ColumnArgument[] ColumnArguments { get; }
        public ClassType ClassType { get; }
        public string SheetName { get; }
    }
    public enum ClassType {
        Default = 1
    }
    public enum ColumnValueType {
        String,
        Int,
        DateTime,
        Double,
        Picture
    }
    public class ColumnArgument {
        public ColumnArgument(int width, ColumnValueType columnValueType) {
            this.Width = width * 170;
            this.ColumnValueType = columnValueType;
        }
        public int Width { get; }
        public ColumnValueType ColumnValueType { get; }
    }
}
