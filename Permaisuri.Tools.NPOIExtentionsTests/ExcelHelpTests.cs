using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Threading;

namespace Permaisuri.Tools.NPOIExtentions.Tests {
    [TestClass()]
    public class ExcelHelpTests {
        private string outPutDirPath = "../../OutPutDir/" + DateTime.Now.Millisecond.ToString();
        public static string PicPath = "../../Pictures/1.jpg";
        [TestMethod()]
        public void GetXlsByDt() {
            var dt = getDt();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();

            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet", height: 6);
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDT.xls"), sheetsArgs);
            new ExcelHelp().GetXls(dt, excelArgs);
        }
        public void GetXlsByDs() {
            var ds = getDs();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet1", height: 6);

            Argument.ColumnArgument[] colArgs2 = getNormalColArgs();
            Argument.SheetArgument sheetArgs2 = new Argument.SheetArgument(colArgs, "TestSheet2", height: 2);

            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs, sheetArgs2 };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDS.xls"), sheetsArgs);
            new ExcelHelp().GetXls(ds, excelArgs);
        }
        public void GetXlsDynamic() {
            var d = getDynamic();
            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet", height: 1);
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDynamic.xls"), sheetsArgs);
            new ExcelHelp().GetXls(d, excelArgs);
        }
        public void GetXlsByDynamicList() {
            var ds = getDynamics();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet1", height: 6);

            Argument.ColumnArgument[] colArgs2 = getNormalColArgs();
            Argument.SheetArgument sheetArgs2 = new Argument.SheetArgument(colArgs, "TestSheet2", height: 2);

            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs, sheetArgs2 };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDynamics.xls"), sheetsArgs);
            new ExcelHelp().GetXls(ds, excelArgs);
        }
        public void GetXlsxByDt() {
            var dt = getDt();
            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet");
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDT.xlsx"), sheetsArgs);
            new ExcelHelp().GetXlsx(dt, excelArgs);
        }
        public void GetXlsxByDs() {
            var ds = getDs();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet1", height: 6);

            Argument.ColumnArgument[] colArgs2 = getNormalColArgs();
            Argument.SheetArgument sheetArgs2 = new Argument.SheetArgument(colArgs, "TestSheet2", height: 2);

            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs, sheetArgs2 };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDS.xlsx"), sheetsArgs);
            new ExcelHelp().GetXlsx(ds, excelArgs);
        }
        public void GetXlsxDynamic() {
            var d = getDynamic();
            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet", height: 6);
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDynamic.xlsx"), sheetsArgs);
            new ExcelHelp().GetXlsx(d, excelArgs);
        }
        public void GetXlsxByDynamicList() {
            var ds = getDynamics();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet1", height: 6);

            Argument.ColumnArgument[] colArgs2 = getNormalColArgs();
            Argument.SheetArgument sheetArgs2 = new Argument.SheetArgument(colArgs, "TestSheet2", height: 2);

            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs, sheetArgs2 };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDynamics.xlsx"), sheetsArgs);
            new ExcelHelp().GetXlsx(ds, excelArgs);
        }
        private DataTable getDt() {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Text");
            dt.Columns.Add("Datetime");
            dt.Columns.Add("DoubleValue");
            dt.Columns.Add("Pictures");

            for (int i = 0; i < 100; i++) {
                DataRow nRow = dt.NewRow();
                nRow["ID"] = i;
                nRow["Text"] = "123123123" + i;
                nRow["Datetime"] = DateTime.Now.AddDays(i);
                nRow["DoubleValue"] = new Random().NextDouble();
                if (i % 2 == 0) {
                    nRow["Pictures"] = "../../Pictures/1.jpg";
                }
                else {
                    nRow["Pictures"] = "../../Pictures/nopic.jpg";
                }
                dt.Rows.Add(nRow);
            }
            ExcelHelp eh = new ExcelHelp();
            return dt;
        }
        private DataSet getDs() {
            DataTable dt = getDt();
            DataTable dt1 = getDt();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.Tables.Add(dt1);
            return ds;
        }
        private dynamic getDynamic() {
            List<User> user = new List<User>();
            for (int i = 0; i < 100; ++i) {
                user.Add(new User(true));
            }
            dynamic d = user.Select(a => new {
                Age = a.Age,
                Name = a.Name,
                BirthDay = a.BirthDate,
                Height = a.Height,
                Pic = a.PicturePath
            }).ToList();

            return d;
        }
        private List<dynamic> getDynamics() {
            List<dynamic> dList = new List<dynamic>();
            dList.Add(getDynamic());
            dList.Add(getDynamic());
            return dList;
        }
        private Argument.ColumnArgument[] getNormalColArgs() {
            Argument.ColumnArgument[] colArgs = new Argument.ColumnArgument[] {
            new Argument.ColumnArgument( 10, Argument.ColumnValueType.Int),
            new Argument.ColumnArgument(20, Argument.ColumnValueType.String),
            new Argument.ColumnArgument(30, Argument.ColumnValueType.DateTime),
            new Argument.ColumnArgument(40, Argument.ColumnValueType.Double),
            new Argument.ColumnArgument(12*2, Argument.ColumnValueType.Picture)
            };
            return colArgs;
        }
    }
    public class User {
        public User() {
        }
        public User(bool isDefalt) {
            if (!isDefalt) { }
            else {
                Name = "Lilei" + new Random().Next(1, 10000);
                Age = new Random().Next(10, 50);
                Height = 182.25;
                BirthDate = DateTime.Now.AddDays(0 - new Random().Next(365 * 10, 365 * 100));
                PicturePath = ExcelHelpTests.PicPath;
                Thread.Sleep(10);
            }
        }
        public string Name { get; set; }
        public int Age { get; set; }
        public string PicturePath { get; set; }
        public double Height { get; set; }
        public DateTime BirthDate { get; set; }
    }
}