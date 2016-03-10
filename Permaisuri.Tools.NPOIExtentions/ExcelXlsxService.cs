using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Permaisuri.Tools.NPOIExtentions.Argument;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace Permaisuri.Tools.NPOIExtentions {
    internal class ExcelXlsxService {
        public static void Export(DataTable dt, ExcelArgument arg) {
            XSSFWorkbook hssfworkbook = createWorkbook();
            createSheet(hssfworkbook, dt, arg.SheetArguments[0]);
            ExcelCommonService.WriteToFile(hssfworkbook, arg.OutPutPath);
        }
        public static void Export(DataSet ds, ExcelArgument arg) {
            XSSFWorkbook hssfworkbook = createWorkbook();
            for (int i = 0; i < ds.Tables.Count; i++) {
                createSheet(hssfworkbook, ds.Tables[i], arg.SheetArguments[i]);
            }
            ExcelCommonService.WriteToFile(hssfworkbook, arg.OutPutPath);
        }
        public static void Export(dynamic list, ExcelArgument arg) {
            XSSFWorkbook hssfworkbook = createWorkbook();
            createSheet(hssfworkbook, list, arg.SheetArguments[0]);
            ExcelCommonService.WriteToFile(hssfworkbook, arg.OutPutPath);
        }
        public static void Export(List<dynamic> lists, ExcelArgument arg) {
            XSSFWorkbook hssfworkbook = createWorkbook();
            for (int i = 0; i < lists.Count; i++) {
                createSheet(hssfworkbook, lists[i], arg.SheetArguments[i]);
            }
            ExcelCommonService.WriteToFile(hssfworkbook, arg.OutPutPath);
        }
        private static XSSFWorkbook createWorkbook() {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            return xssfworkbook;
        }
        private static void createSheet(XSSFWorkbook xssfworkbook, DataTable dt, SheetArgument arg) {
            ISheet sheet1 = xssfworkbook.CreateSheet(arg.SheetName);
            int colCount = dt.Columns.Count;
            int rowCount = dt.Rows.Count;
            int rowNumber = 0;
            IRow rowTitle = sheet1.CreateRow(rowNumber++);
            rowTitle.Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                StyleProcesser.SetSheetDefaultStyle(arg.ColumnArguments[i].ColumnValueType, xssfworkbook, sheet1, i);
                ICell cell = rowTitle.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
                cell.CellStyle = StyleProcesser.GetStyleByClass(arg.ClassType, xssfworkbook);
                sheet1.SetColumnWidth(i, arg.ColumnArguments[i].Width);
            }
            for (int i = 0; i < rowCount; ++i) {
                DataRow dr = dt.Rows[i];
                IRow rowContent = sheet1.CreateRow(rowNumber++);
                rowContent.Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    setCellValue(dr[ii], arg.ColumnArguments[ii].ColumnValueType, rowContent.CreateCell(ii), xssfworkbook, sheet1);
                }
            }
        }
        private static void createSheet(XSSFWorkbook xssfworkbook, dynamic list, SheetArgument arg) {
            ISheet sheet1 = xssfworkbook.CreateSheet(arg.SheetName);
            Type listType = list.GetType();
            MethodInfo m_Count = listType.GetMethod("Count");
            Type type = null;
            foreach (var l in list) {
                type = l.GetType();
                break;
            }
            var properties = type.GetProperties();
            int colCount = properties.Length;
            int rowNumber = 0;
            IRow rowTitle = sheet1.CreateRow(rowNumber++);
            rowTitle.Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                StyleProcesser.SetSheetDefaultStyle(arg.ColumnArguments[i].ColumnValueType, xssfworkbook, sheet1, i);
                ICell cell = rowTitle.CreateCell(i);
                cell.SetCellValue(properties[i].Name);
                cell.CellStyle = StyleProcesser.GetStyleByClass(arg.ClassType, xssfworkbook);
                sheet1.SetColumnWidth(i, arg.ColumnArguments[i].Width);
            }
            foreach (var l in list) {
                var row = l;
                IRow rowContent = sheet1.CreateRow(rowNumber++);
                rowContent.Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    setCellValue(properties[ii].GetValue(row), arg.ColumnArguments[ii].ColumnValueType, rowContent.CreateCell(ii), xssfworkbook, sheet1);
                }
            }
        }
        private static void setCellValue(object value, ColumnValueType type, ICell cell, XSSFWorkbook xssfworkbook, ISheet sheet1) {
            switch (type) {
                case ColumnValueType.String:
                    cell.SetCellValue(value != DBNull.Value && value != null ? Convert.ToString(value) : "");
                    break;
                case ColumnValueType.Int:
                    cell.SetCellValue(value != DBNull.Value && value != null ? Convert.ToInt32(value) : 0);
                    break;
                case ColumnValueType.Double:
                    cell.SetCellValue(value != DBNull.Value && value != null ? Convert.ToDouble(value) : 0.00);
                    break;
                case ColumnValueType.DateTime:
                    cell.CellStyle = sheet1.GetColumnStyle(cell.ColumnIndex);
                    if (value != DBNull.Value && value != null) {
                        cell.SetCellValue(Convert.ToDateTime(value));
                    }
                    else {
                        cell.SetCellValue("");
                    }
                    break;
                case ColumnValueType.Picture:
                    PictureXlsxProcesser.SetPictureToCell(value.ToString(), cell, sheet1, xssfworkbook);
                    break;
                case ColumnValueType.IntNull:
                    if (value != DBNull.Value && value != null) {
                        cell.SetCellValue(Convert.ToInt32(value));
                    }
                    else {
                        cell.SetCellValue("");
                    }
                    break;
                case ColumnValueType.DoubleNull:
                    if (value != DBNull.Value && value != null) {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else {
                        cell.SetCellValue("");
                    }
                    break;
                default: break;
            }
        }
    }
}
