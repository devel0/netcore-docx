using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;

namespace SearchAThing.DocX
{

    public enum PaperType
    {
        A4
    };

    public static class DocXUtil
    {

        static int MMToTwip(double mm)
        {
            return (int)(mm / 25.4 * 1440);
        }

        /// <summary>
        /// set section pagesize to portrait of given paper type
        /// </summary>
        static void SetPaper(this SectionProperties section, PaperType paper)
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
            pageSize.Width = (uint)MMToTwip(w);
            pageSize.Height = (uint)MMToTwip(h);
            pageSize.Orient = new EnumValue<PageOrientationValues>(PageOrientationValues.Portrait);
        }

        /// <summary>
        /// set section pagesize orientation
        /// </summary>
        static void SetOrientation(this SectionProperties section, PageOrientationValues orientation)
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

        static void SetMargin(this SectionProperties section,
            int marginLeftMM = 0,
            int marginTopMM = 0,
            int marginRightMM = 0,
            int marginBottomMM = 0)
        {
            var margin = section.Descendants<PageMargin>().FirstOrDefault();
            if (margin == null)
                margin = section.AppendChild(new PageMargin());
            margin.Left = (uint)MMToTwip(marginLeftMM);
            margin.Top = MMToTwip(marginTopMM);
            margin.Right = (uint)MMToTwip(marginRightMM);
            margin.Bottom = MMToTwip(marginBottomMM);
        }

        /// <summary>
        /// create new empty doc
        /// </summary>
        public static WordprocessingDocument Create(string pathfilename,
            PaperType paperType = PaperType.A4,
            PageOrientationValues orientation = PageOrientationValues.Portrait,
            int marginLeftMM = 0,
            int marginTopMM = 0,
            int marginRightMM = 0,
            int marginBottomMM = 0)
        {
            using (var doc = WordprocessingDocument.Create(pathfilename, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document();

                var body = main.Document.AppendChild(new Body());

                var p = body.AppendChild(new Paragraph());
                var r = p.AppendChild(new Run());
                r.AppendChild(new Text(DateTime.Now.ToString()));

                var sect = body.AppendChild(new SectionProperties());
                sect.SetPaper(paperType);
                sect.SetOrientation(orientation);
                sect.SetMargin(marginLeftMM, marginTopMM, marginRightMM, marginBottomMM);

                doc.Save();

                return doc;
            }
        }
    }


}
