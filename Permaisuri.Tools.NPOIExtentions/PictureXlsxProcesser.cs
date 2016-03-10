using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Drawing;


namespace Permaisuri.Tools.NPOIExtentions {
    internal class PictureXlsxProcesser {
        private const int scale = 1;
        private static int loadImage(string path, IWorkbook wb, ref Image img) {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            img = Image.FromStream(file);
            return wb.AddPicture(buffer, PictureType.JPEG);
        }
        public static void SetPictureToCell(string picturePath, ICell cell, ISheet sheet1, XSSFWorkbook hssfworkbook) {
            if (File.Exists(picturePath)) {
                IDrawing patriarch = sheet1.CreateDrawingPatriarch();
                //create the anchor
                XSSFClientAnchor anchor;
                anchor = new XSSFClientAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex, cell.RowIndex);
                Image img = null;
                int picIndex = loadImage(picturePath.ToString(), hssfworkbook, ref img);
                IPicture picture = patriarch.CreatePicture(anchor, picIndex);

                var picWidth = img.Width;
                var picHeight = img.Height;

                int cellWidth = sheet1.GetColumnWidth(cell.ColumnIndex) / 17;
                var row = sheet1.GetRow(cell.RowIndex);
                int rowHeight = row.Height / 17;
                int minRowHeight = Convert.ToInt32(picHeight / 2 / 0.8 * scale);
                if (rowHeight < minRowHeight) {
                    rowHeight = minRowHeight;
                    row.Height = Convert.ToInt16(minRowHeight * 17);
                }
                int offsetX = picWidth * 6000 * scale;
                int offsetY = picHeight * 6000 * scale;
                int startX = (cellWidth * 4418 - offsetX) / 2;
                int startY = (rowHeight * 10668 - offsetY) / 2;
                anchor.Dx1 = startX;
                anchor.Dy1 = startY;
                anchor.Dx2 = startX + offsetX;
                anchor.Dy2 = startY + offsetY;
                anchor.AnchorType = 0;
            }
            else {
                cell.SetCellValue("No Pic");
            }
        }

        #region For NPOI 2.1.3.1
        //public static void SetPictureToCell(string picturePath, ICell cell, ISheet sheet1, IWorkbook hssfworkbook) {
        //    if (File.Exists(picturePath)) {
        //        IDrawing patriarch = sheet1.CreateDrawingPatriarch();
        //        //create the anchor
        //        XSSFClientAnchor anchor;
        //        anchor = new XSSFClientAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex, cell.RowIndex);
        //        Image img = null;
        //        int picIndex = loadImage(picturePath.ToString(), hssfworkbook, ref img);
        //        IPicture picture = patriarch.CreatePicture(anchor, picIndex);

        //        var picWidth = img.Width;
        //        var picHeight = img.Height;

        //        int cellWidth = sheet1.GetColumnWidth(cell.ColumnIndex) / 17;
        //        var row = sheet1.GetRow(cell.RowIndex);
        //        int rowHeight = row.Height / 17;
        //        int minRowHeight = Convert.ToInt32(picHeight / 0.9);
        //        if (rowHeight < minRowHeight) {
        //            rowHeight = minRowHeight;
        //            row.Height = Convert.ToInt16(minRowHeight * 17);
        //        }
        //        var preAnchor = picture.GetPreferredSize();
        //        int offsetX = preAnchor.Dx2;
        //        int offsetY = preAnchor.Dy2;
        //        int startX = 0;
        //        if (preAnchor.Col1 == preAnchor.Col2) {
        //            startX = (cellWidth / 2 * 9525 - offsetX) / 2;
        //        }
        //        else {
        //            startX = 0;
        //        }
        //        int startY = (rowHeight * 9525 - offsetY) / 2;
        //        anchor.Dx1 = startX;
        //        anchor.Dy1 = startY;
        //        anchor.Dx2 = startX + offsetX;
        //        anchor.Dy2 = startY + offsetY;
        //        anchor.AnchorType = 0;
        //    }
        //    else {
        //        cell.SetCellValue("No Pic");
        //    }
        //}
        #endregion
    }
}
