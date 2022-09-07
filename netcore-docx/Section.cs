using System;
using System.Linq;
using static System.Math;
using System.Collections.Generic;
using System.Globalization;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using SearchAThing;
using SearchAThing.DocX;
using static SearchAThing.DocX.Constants;
using System.Runtime.CompilerServices;

namespace SearchAThing.DocX
{

    public static partial class DocXExt
    {

        /// <summary>
        /// create Header element in the document; this could set section of paragraphs until
        /// the one where set .SectionProperties().SetHeader()
        /// </summary>
        public static Header CreateHeader(this WordprocessingDocument doc)
        {
            var part = doc.CreateHeaderPart();

            var header = new Header();

            part.Header = header;

            return header;
        }

        /// <summary>
        /// create Footer element in the document; this could set section of paragraphs until
        /// the one where set .SectionProperties().SetFooter()
        /// </summary>
        public static Footer CreateFooter(this WordprocessingDocument doc)
        {
            var part = doc.CreateFooterPart();

            var footer = new Footer();

            part.Footer = footer;

            return footer;
        }

        /// <summary>
        /// set the main section properties.
        /// see also: Paragraph.SectionProperties()
        /// </summary>        
        public static SectionProperties MainSectionProperties(this WordprocessingDocument doc) =>
            doc.GetMainSectionProperties(createIfNotExists: true)!;

        internal static SectionProperties? GetMainSectionProperties(this WordprocessingDocument doc, bool createIfNotExists = false, Body? _body = null)
        {
            var body = _body is null ? doc.GetBody() : _body;
            var mainSectionProperties = body.Elements<SectionProperties>().FirstOrDefault();

            if (mainSectionProperties is null && createIfNotExists)
            {
                mainSectionProperties = new SectionProperties();
                body.Append(mainSectionProperties);
            }

            return mainSectionProperties;
        }

        /// <summary>
        /// retrieve current section properties.
        /// it search for a section properties from current paragraph to latest;
        /// if not found main document section properties will returned
        /// </summary>
        /// <param name="paragraph">paragraph from which starts search</param>
        /// <returns>section properties associated with current paragraph</returns>
        public static SectionProperties CurrentSectionProperties(this Paragraph paragraph)
        {
            var doc = paragraph.GetWordprocessingDocument();

            var body = doc.GetBody();

            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true)!;

            if (paragraph.HasSectionProperties())
                return paragraphProperties.SectionProperties!;

            var nextParagraphWithSectionProperties = paragraph
                .NextSibling<Paragraph>()?
                .OfType<Paragraph>()
                .Cast<Paragraph>()
                .Where(paragraph => paragraph.HasSectionProperties())
                .FirstOrDefault();

            SectionProperties sectionProperties;

            if (nextParagraphWithSectionProperties is null) // next section properties are main section properties
            {
                sectionProperties = doc.GetMainSectionProperties(_body: body)!;
            }
            else // next section properties are from another paragraph
            {
                sectionProperties = nextParagraphWithSectionProperties.GetProperties()!.SectionProperties!;
            }

            return sectionProperties;
        }

        /// <summary>
        /// retrieve current section width (mm)
        /// </summary>
        /// <param name="paragraph">paragraph from where starts section properties search</param>
        /// <returns>current section width (mm)</returns>
        public static double GetCurrentSectionWidthMM(this Paragraph paragraph) =>
            paragraph.CurrentSectionProperties().GetPageSize()!.Width!.TwipToMM();

        /// <summary>
        /// if this paragraph has already a section properties that will returned,
        /// else a new section property will returned.
        /// 
        /// note: if there isnt' a section property already defined
        /// in the paragraph itself, a section property will searched in next paragraph and if there isn't one
        /// then main section properties will used for the new created paragraph which want to set section properties,
        /// while a new clean main section property will assigned to the document body.
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="insertAtIdx"></param>
        /// <returns></returns>
        public static SectionProperties SectionProperties(this Paragraph paragraph, int? insertAtIdx = null)
        {
            var doc = paragraph.GetWordprocessingDocument();

            var body = doc.GetBody();

            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true)!;

            if (paragraph.HasSectionProperties())
                return paragraphProperties.SectionProperties!;

            var nextParagraphWithSectionProperties = paragraph
                .NextSibling<Paragraph>()?
                .Cast<Paragraph>()
                .Where(paragraph => paragraph.HasSectionProperties())
                .FirstOrDefault();

            var sectionProperties = new SectionProperties();

            if (nextParagraphWithSectionProperties is null) // next section properties are main section properties
            {
                var mainSectionProperties = doc.GetMainSectionProperties(_body: body)!;
                mainSectionProperties.Remove();
                paragraphProperties.Append(mainSectionProperties); // main section properties moved to this paragraph

                sectionProperties.CopyFrom(mainSectionProperties);
                body.Append(sectionProperties); // new main section properties
            }
            else // next section properties are from another paragraph
            {
                var nextParagraphProperties = nextParagraphWithSectionProperties.GetProperties()!;
                var nextSectionProperties = nextParagraphProperties.SectionProperties!;
                nextSectionProperties.Remove();
                paragraphProperties.SectionProperties = nextSectionProperties; // next paragraph section prop move to this paragraph

                nextParagraphProperties.Append(sectionProperties);
            }

            return sectionProperties;
        }

        /// <summary>
        /// states if paragraph has a Section Properties in its paragraph properties
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool HasSectionProperties(this Paragraph paragraph) => paragraph.GetProperties()?.SectionProperties is not null;

        /// <summary>
        /// states if document has a Section Properties in its body
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static bool HasSectionProperties(this WordprocessingDocument doc) => doc.GetMainSectionProperties() is not null;

        public static PageSize? GetPageSize(this SectionProperties sectionProperties) =>
            sectionProperties.Elements<PageSize>().FirstOrDefault();

        public static SectionProperties SetPageSize(this SectionProperties sectionProperties,
            PaperSize size,
            PageOrientationValues? orientation = null)
        {
            var w = (uint)size.WidthMM.MMToTwip();
            var h = (uint)size.HeightMM.MMToTwip();
            if (orientation is not null && orientation.Value == PageOrientationValues.Landscape)
                UtilToolkit.Swap(ref w, ref h);

            var pageSize = new PageSize
            {
                Width = w,
                Height = h,
                Orient = orientation
            };
            sectionProperties.Append(pageSize);

            return sectionProperties;
        }

        public static SectionProperties SetMargin(this SectionProperties sectionProperties,
            double? marginLeftMM = null, double? marginTopMM = null, double? marginRightMM = null, double? marginBottomMM = null)
        {
            var margin = sectionProperties.Elements<PageMargin>().FirstOrDefault();
            if (margin is null)
            {
                margin = new PageMargin();
                sectionProperties.Append(margin);
            }

            margin.Left = (uint?)marginLeftMM?.MMToTwip();
            margin.Top = (int?)marginTopMM?.MMToTwip();
            margin.Right = (uint?)marginRightMM?.MMToTwip();
            margin.Bottom = (int?)marginBottomMM?.MMToTwip();

            return sectionProperties;
        }

        public static SectionProperties SetHeader(this SectionProperties sectionProperties, Action<Header> action,
            HeaderFooterValues type = HeaderFooterValues.Default)
        {
            var doc = sectionProperties.GetWordprocessingDocument();

            sectionProperties.SetHeader(doc.CreateHeader().Act(header => action(header)));

            return sectionProperties;
        }

        internal static SectionProperties SetHeader(this SectionProperties sectionProperties, Header header,
            HeaderFooterValues type = HeaderFooterValues.Default)
        {
            var id = header.GetMainDocumentPart().GetIdOfPart(header.HeaderPart!);

            {
                var toremove = sectionProperties.Elements<HeaderReference>().ToList();
                foreach (var x in toremove) x.Remove();
            }

            sectionProperties.InsertAt(new HeaderReference { Type = type, Id = id }, 0);

            return sectionProperties;
        }

        public static SectionProperties SetFooter(this SectionProperties sectionProperties, Action<Footer> action,
            HeaderFooterValues type = HeaderFooterValues.Default)
        {
            var doc = sectionProperties.GetWordprocessingDocument();

            sectionProperties.SetFooter(doc.CreateFooter().Act(footer => action(footer)));

            return sectionProperties;
        }

        public static SectionProperties SetFooter(this SectionProperties sectionProperties, Footer footer,
            HeaderFooterValues type = HeaderFooterValues.Default)
        {
            var id = footer.GetMainDocumentPart().GetIdOfPart(footer.FooterPart!);

            {
                var toremove = sectionProperties.Elements<FooterReference>().ToList();
                foreach (var x in toremove) x.Remove();
            }

            sectionProperties.InsertAt(new FooterReference { Type = type, Id = id }, 0);

            return sectionProperties;
        }

    }

}