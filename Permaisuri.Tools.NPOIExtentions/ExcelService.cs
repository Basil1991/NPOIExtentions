﻿using Permaisuri.Tools.NPOIExtentions.Argument;
using System;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections.Generic;
using System.Reflection;

namespace Permaisuri.Tools.NPOIExtentions {
    internal static class ExcelXlsService {
        public static void ExportXlsExcel(DataTable dt, ExcelArgument arg) {
            HSSFWorkbook hssfworkbook = createWorkbook();
            createSheet(hssfworkbook, dt, arg.SheetArguments[0]);
            writeToFile(hssfworkbook, arg.OutPutPath);
        }
        public static void ExportXlsExcel(DataSet ds, ExcelArgument arg) {
            HSSFWorkbook hssfworkbook = createWorkbook();
            for (int i = 0; i < ds.Tables.Count; i++) {
                createSheet(hssfworkbook, ds.Tables[i], arg.SheetArguments[i]);
            }
            writeToFile(hssfworkbook, arg.OutPutPath);
        }
        public static void ExportXlsExcel(dynamic list, ExcelArgument arg) {
            HSSFWorkbook hssfworkbook = createWorkbook();
            createSheet(hssfworkbook, list, arg.SheetArguments[0]);
            writeToFile(hssfworkbook, arg.OutPutPath);
        }
        public static void ExportXlsExcel(List<dynamic> lists, ExcelArgument arg) {
            HSSFWorkbook hssfworkbook = createWorkbook();
            for (int i = 0; i < lists.Count; i++) {
                createSheet(hssfworkbook, lists[i], arg.SheetArguments[i]);
            }
            writeToFile(hssfworkbook, arg.OutPutPath);
        }
        private static void createSheet(HSSFWorkbook hssfworkbook, DataTable dt, SheetArgument arg) {
            ISheet sheet1 = hssfworkbook.CreateSheet(arg.SheetName);
            int colCount = dt.Columns.Count;
            int rowNumber = 0;
            IRow rowTitle = sheet1.CreateRow(rowNumber++);
            rowTitle.Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                ICell cell = rowTitle.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
                cell.CellStyle = StyleProcesser.GetStyleByClass(arg.ClassType, hssfworkbook);
                sheet1.SetColumnWidth(i, arg.ColumnArguments[i].Width);
            }
            for (int i = 0; i < dt.Rows.Count; ++i) {
                DataRow dr = dt.Rows[i];
                IRow rowContent = sheet1.CreateRow(rowNumber++);
                rowContent.Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    setCellValue(dr[ii], arg.ColumnArguments[ii].ColumnValueType, rowContent.CreateCell(ii), hssfworkbook, sheet1);
                }
            }
        }
        private static void createSheet(HSSFWorkbook hssfworkbook, dynamic list, SheetArgument arg) {
            ISheet sheet1 = hssfworkbook.CreateSheet(arg.SheetName);
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
                ICell cell = rowTitle.CreateCell(i);
                cell.SetCellValue(properties[i].Name);
                cell.CellStyle = StyleProcesser.GetStyleByClass(arg.ClassType, hssfworkbook);
                sheet1.SetColumnWidth(i, arg.ColumnArguments[i].Width);
            }
            foreach (var l in list) {
                var row = l;
                IRow rowContent = sheet1.CreateRow(rowNumber++);
                rowContent.Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    setCellValue(properties[ii].GetValue(row), arg.ColumnArguments[ii].ColumnValueType, rowContent.CreateCell(ii), hssfworkbook, sheet1);
                }
            }
        }
        private static HSSFWorkbook createWorkbook() {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            //////create a entry of DocumentSummaryInformation
            //DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            //dsi.Company = "Permaisuri";
            //hssfworkbook.DocumentSummaryInformation = dsi;

            //////create a entry of SummaryInformation
            //SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            //si.Subject = "Excel";
            //hssfworkbook.SummaryInformation = si;
            return
                 hssfworkbook;
        }
        private static void setCellValue(object value, ColumnValueType type, ICell cell, HSSFWorkbook hssfworkbook, ISheet sheet1) {
            switch (type) {
                case ColumnValueType.String:
                    cell.SetCellValue(Convert.ToString(value));
                    break;
                case ColumnValueType.Int:
                    cell.SetCellValue(Convert.ToInt32(value));
                    break;
                case ColumnValueType.Double:
                    cell.SetCellValue(Convert.ToDouble(value));
                    break;
                case ColumnValueType.DateTime:
                    cell.SetCellValue(Convert.ToDateTime(value));
                    cell.CellStyle = StyleProcesser.GetCellStyleByDataType(type, hssfworkbook);
                    break;
                case ColumnValueType.Picture:
                    PictureProcesser.SetPictureToCell(value.ToString(), cell, sheet1, hssfworkbook);
                    break;
                default: break;
            }
        }
        private static void writeToFile(HSSFWorkbook hssfworkbook, string outPutPath) {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(outPutPath, FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }
    }
}
