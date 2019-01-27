using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;
using SearchAThing.DocX;

namespace SearchAThing.DocX
{

    public enum PaperType
    {
        A4
    };

    public static class SectionExt
    {

        /// <summary>
        /// set section pagesize to portrait of given paper type
        /// </summary>
        public static void SetPaper(this SectionProperties section, PaperType paper)
        {
            var pageSize = section.Descendants<PageSize>().FirstOrDefault();
            if (pageSize == null)
                pageSize = section.AppendChild(new PageSize());
            int w = 0;
            int h = 0;
            switch (paper)
            {
                case PaperType.A4:
                    {
                        w = 210;
                        h = 297;
                    }
                    break;
                default: throw new Exception($"unsupported paper type {paper}");
            }
            pageSize.Width = (uint)w.MMToTwip();
            pageSize.Height = (uint)h.MMToTwip();
            pageSize.Orient = new EnumValue<PageOrientationValues>(PageOrientationValues.Portrait);
        }

        /// <summary>
        /// retrieve section page size info
        /// </summary>
        public static (double widthMM, double heightMM, PageOrientationValues orient) GetPageSize(this SectionProperties section)
        {
            var pageSize = section.Descendants<PageSize>().First();
            return (Round(pageSize.Width.Value.TwipToMM(), 0), Round(pageSize.Height.Value.TwipToMM(), 0), pageSize.Orient.Value);
        }

        /// <summary>
        /// set section pagesize orientation
        /// </summary>
        public static void SetOrientation(this SectionProperties section, PageOrientationValues orientation)
        {
            var pageSize = section.Descendants<PageSize>().FirstOrDefault();
            if (pageSize == null) throw new Exception($"must set paper first");
            if (pageSize.Orient.Value != orientation)
            {
                var w = pageSize.Width;
                var h = pageSize.Height;
                pageSize.Width = h;
                pageSize.Height = w;
                pageSize.Orient = new EnumValue<PageOrientationValues>(orientation);

                var margin = section.Descendants<PageMargin>().FirstOrDefault();
                if (margin != null)
                {
                    var left = margin.Left.Value;
                    var top = margin.Top.Value;
                    var right = margin.Right.Value;
                    var bottom = margin.Bottom.Value;

                    margin.Top = (int)left;
                    margin.Bottom = (int)right;
                    margin.Left = (uint)Max(0, bottom);
                    margin.Right = (uint)Max(0, top);
                }
            }
        }

        /// <summary>
        /// set section margin
        /// </summary>
        public static void SetMargin(this SectionProperties section,
            double marginLeftMM = 0,
            double marginTopMM = 0,
            double marginRightMM = 0,
            double marginBottomMM = 0)
        {
            var margin = section.Descendants<PageMargin>().FirstOrDefault();
            if (margin == null)
                margin = section.AppendChild(new PageMargin());
            margin.Left = (uint)marginLeftMM.MMToTwip();
            margin.Top = marginTopMM.MMToTwip();
            margin.Right = (uint)marginRightMM.MMToTwip();
            margin.Bottom = marginBottomMM.MMToTwip();
        }

    }


}
