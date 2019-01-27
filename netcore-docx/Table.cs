using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;
using SearchAThing.DocX;
using static System.FormattableString;

namespace SearchAThing.DocX
{

    public static class TableExt
    {

        /// <summary>
        /// create a table
        /// </summary>
        public static Table AddTable(this WordprocessingDocument doc, int rowCount, int colCount)
        {
            var tbl = new Table();
            var tblProp = tbl.AppendChild(new TableProperties());

            var docSize = doc.GetMainSection().GetPageSize();
            var pageMargin = doc.GetMainSection().GetPageMargin();
            var colSizeMM = (docSize.widthMM - pageMargin.leftMM - pageMargin.rightMM) / colCount;

            for (int r = 0; r < rowCount; ++r)
            {
                var tr = new TableRow();

                for (int c = 0; c < colCount; ++c)
                {
                    var tc = new TableCell();
                    var tcwidthpc = (100d / colCount).Pct();
                    if (c == colCount - 1)
                        tcwidthpc = 5000 - (100d / colCount).Pct() * colCount - 1;
                    tc.Append(new TableCellProperties(
                       new TableCellWidth()
                       {
                           Type = TableWidthUnitValues.Dxa,
                           Width = colSizeMM.MMToTwip().ToString()
                       }));
                    tr.AppendChild(tc);
                }

                tbl.AppendChild(tr);
            }

            var body = doc.Body();
            return body.AppendChild(tbl);
        }

        /// <summary>
        /// retrieve table cell ref (first row/col = 1)
        /// </summary>
        public static TableCell GetCell(this Table tbl, int row, int col)
        {
            return tbl.GetRow(row).DescendantAt<TableCell>(col - 1);
        }

        /// <summary>
        /// retrieve table row ref (first row/col = 1)
        /// </summary>
        public static TableRow GetRow(this Table tbl, int row)
        {
            return tbl.DescendantAt<TableRow>(row - 1);
        }

        /// <summary>
        /// set row height
        /// </summary>
        public static void SetHeight(this TableRow row, double valMM, HeightRuleValues type = HeightRuleValues.Exact)
        {
            var trProp = row.Descendants<TableRowProperties>().FirstOrDefault();
            if (trProp == null) trProp = row.InsertAt(new TableRowProperties(), 0);
            var trHeight = trProp.Descendants<TableRowHeight>().FirstOrDefault();
            if (trHeight == null) trHeight = trProp.AppendChild(new TableRowHeight());
            trHeight.HeightType = type;
            if (type != HeightRuleValues.Auto)
                trHeight.Val = (uint)valMM.MMToTwip();
        }

        /// <summary>
        /// set cell margin (null values are left unchanged)
        /// </summary>
        public static void SetMargin(this TableCell cell, double? marginLeftMM = null,
            double? marginTopMM = null,
            double? marginRightMM = null,
            double? marginBottomMM = null)
        {
            var tcPr = cell.Descendants<TableCellProperties>().FirstOrDefault();
            if (tcPr == null) tcPr = cell.AppendChild(new TableCellProperties());
            var tcMar = tcPr.Descendants<TableCellMargin>().FirstOrDefault();
            if (tcMar == null) tcMar = tcPr.AppendChild(new TableCellMargin());
            if (marginLeftMM.HasValue)
            {
                if (tcMar.LeftMargin == null)
                    tcMar.LeftMargin = new LeftMargin();
                tcMar.LeftMargin.Width = marginLeftMM.Value.MMToTwip().ToString();
                tcMar.LeftMargin.Type = TableWidthUnitValues.Dxa;
            }
            if (marginTopMM.HasValue)
            {
                if (tcMar.TopMargin == null)
                    tcMar.TopMargin = new TopMargin();
                tcMar.TopMargin.Width = marginTopMM.Value.MMToTwip().ToString();
                tcMar.TopMargin.Type = TableWidthUnitValues.Dxa;
            }
            if (marginRightMM.HasValue)
            {
                if (tcMar.RightMargin == null)
                    tcMar.RightMargin = new RightMargin();
                tcMar.RightMargin.Width = marginRightMM.Value.MMToTwip().ToString();
                tcMar.RightMargin.Type = TableWidthUnitValues.Dxa;
            }
            if (marginBottomMM.HasValue)
            {
                if (tcMar.BottomMargin == null)
                    tcMar.BottomMargin = new BottomMargin();
                tcMar.BottomMargin.Width = marginBottomMM.Value.MMToTwip().ToString();
                tcMar.BottomMargin.Type = TableWidthUnitValues.Dxa;
            }
        }

    }


}
