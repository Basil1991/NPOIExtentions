using Microsoft.VisualStudio.TestTools.UnitTesting;
using Permaisuri.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Permaisuri.Tools.NPOIExtentions.Tests {
    [TestClass()]
    public class ExcelHelpTests {
        [TestMethod()]
        public void GetTest() {
            var dt = getTestDt();

            Argument.ColumnArgument[] colArgs = new Argument.ColumnArgument[] {
            new Argument.ColumnArgument( 10, Argument.ColumnValueType.Int),
            new Argument.ColumnArgument(20, Argument.ColumnValueType.String),
            new Argument.ColumnArgument(30, Argument.ColumnValueType.DateTime),
            new Argument.ColumnArgument(40, Argument.ColumnValueType.Double),
            new Argument.ColumnArgument(12*3, Argument.ColumnValueType.Picture)
            };

            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet", height: 6);
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(@"C:\Test\{0}.xls", DateTime.Now.Millisecond), sheetsArgs);
            new ExcelHelp().Get(dt, excelArgs);
        }
        private DataTable getTestDt() {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Text");
            dt.Columns.Add("Datetime");
            dt.Columns.Add("DoubleValue");
            dt.Columns.Add("Pictures");

            for (int i = 0; i < 10; i++) {
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
        private DataSet getTestDs() {
            DataTable dt = getTestDt();
            DataTable dt1 = getTestDt();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.Tables.Add(dt);
            return ds;
        }
    }
}