using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.IO;

namespace Permaisuri.Tools.NPOIExtentions {
    internal class PictureXlsProcesser {
        private const int scale = 1;
        private static int loadImage(string path, IWorkbook wb) {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            return wb.AddPicture(buffer, PictureType.JPEG);
        }
        public static void SetPictureToCell(string picturePath, ICell cell, ISheet sheet1, HSSFWorkbook hssfworkbook) {
            if (File.Exists(picturePath)) {
                HSSFPatriarch patriarch = (HSSFPatriarch)sheet1.CreateDrawingPatriarch();
                //create the anchor
                HSSFClientAnchor anchor;
                anchor = new HSSFClientAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex, cell.RowIndex);
                HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, loadImage(picturePath.ToString(), hssfworkbook));

                var picSize = picture.GetImageDimension();
                //var picWidth = picSize.Width;
                var picHeight = picSize.Height;

                int cellWidth = sheet1.GetColumnWidth(cell.ColumnIndex) / 17;
                var row = sheet1.GetRow(cell.RowIndex);
                int rowHeight = row.Height / 17;
                int minRowHeight = Convert.ToInt32(picHeight / 2 / 0.8 * scale);
                if (rowHeight < minRowHeight) {
                    rowHeight = minRowHeight;
                    row.Height = Convert.ToInt16(minRowHeight * 17);
                }
                picture.Resize();
                int offsetX = anchor.Dx2 * scale;
                int offsetY = anchor.Dy2 * scale;
                int startX = 0;
                if (anchor.Col1 == anchor.Col2) {
                    startX = (1024 - offsetX) / 2;
                    startX = startX < 0 ? 0 : startX;
                }
                int startY = (256 - offsetY) / 2;
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
        //        HSSFPatriarch patriarch = (HSSFPatriarch)sheet1.CreateDrawingPatriarch();
        //        //create the anchor
        //        HSSFClientAnchor anchor;
        //        anchor = new HSSFClientAnchor(0, 0, 0, 0, cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex, cell.RowIndex);
        //        HSSFPicture picture = (HSSFPicture)patriarch.CreatePicture(anchor, loadImage(picturePath.ToString(), hssfworkbook));

        //        var picSize = picture.GetImageDimension();
        //        //var picWidth = picSize.Width;
        //        var picHeight = picSize.Height;

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

        //        int startX = (1024 - offsetX) / 2;
        //        int startY = (256 - offsetY) / 2;
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
