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

    public enum LibraryStyleEnum
    {
        Normal,

        Heading,
        Heading1,
        Heading2,
        Heading3,
        Heading4,
        Heading5,
        Heading6,
        Heading7,
        Heading8,
        Heading9,

        InternetLink,
        Title,
        Subtitle,
        PreformattedText,

        Figure,
        Table,
        Drawing,

        Index,
        Contents1,
        IndexHeading,
        ContentsHeading
    }

    public partial class WrapperManager
    {                 

        /// <summary>
        /// scan element for ParagraphStyleId, RunStyle used and integrate if missing referenced styles
        /// </summary>
        /// <param name="element">element to scan</param>        
        /// <returns>element</returns>
        public T IntegrateRequiredStyles<T>(T element) where T : OpenXmlElement
        {
            var doc = element.GetWordprocessingDocument();

            var lib = GetWrapperRef(doc);

            foreach (var paragraphStyleId in element.Descendants<ParagraphStyleId>())
            {
                if (paragraphStyleId.Val is not null && paragraphStyleId.Val.HasValue)
                {
                    var styleId = paragraphStyleId.Val.Value!;

                    if (!doc.HasStyleId(styleId)) CopyStyle(doc, styleId);
                }
            }

            foreach (var runStyle in element.Descendants<RunStyle>())
            {
                if (runStyle.Val is not null && runStyle.Val.HasValue)
                {
                    var styleId = runStyle.Val.Value!;

                    if (!doc.HasStyleId(styleId)) CopyStyle(doc, styleId);
                }
            }

            return element;
        }

        internal void CopyStyle(WordprocessingDocument doc, string styleId)
        {
            if (Enum.TryParse<LibraryStyleEnum>(styleId, ignoreCase: false, out var libraryStyleEnum))
            {
                if (StyleIdToStyleDict.TryGetValue(libraryStyleEnum, out var stylefn))
                {
                    var style = stylefn();
                    var existingStyle = doc.GetStyleById(styleId);
                    var styles = doc.GetStyles(createIfNotExists: true)!;
                    if (existingStyle is not null) styles.RemoveChild(existingStyle);

                    styles.Append(style);
                }
            }
        }

        /// <summary>
        /// tries to find specified styleid, if found style returned but not yet added to document styles
        /// </summary>
        /// <param name="styleId">style id</param>
        /// <returns>style from library</returns>
        internal Style? TryGetStyle(string styleId)
        {
            if (Enum.TryParse<LibraryStyleEnum>(styleId, ignoreCase: false, out var libraryStyleEnum))
            {
                if (StyleIdToStyleDict.TryGetValue(libraryStyleEnum, out var stylefn))
                    return stylefn();
            }

            return null;
        }

        Dictionary<LibraryStyleEnum, Func<Style>>? _StyleIdToStyleDict = null;
        internal Dictionary<LibraryStyleEnum, Func<Style>> StyleIdToStyleDict
        {
            get
            {
                if (_StyleIdToStyleDict is null)
                {
                    var dict = new Dictionary<LibraryStyleEnum, Func<Style>>();

                    var headingAbstractNum_none = doc.GetAbstractNum(NumberFormatValues.None, restartNumbering: true);
                    var headingNumberingInstance_none = doc.GetNumberingInstance(headingAbstractNum_none);

                    dict.Add(LibraryStyleEnum.Normal, () => GenerateStyle_Normal());

                    dict.Add(LibraryStyleEnum.InternetLink, () => GenerateStyle_InternetLink());
                    dict.Add(LibraryStyleEnum.Title, () => GenerateStyle_Title());
                    dict.Add(LibraryStyleEnum.Subtitle, () => GenerateStyle_Subtitle());
                    dict.Add(LibraryStyleEnum.PreformattedText, () => GenerateStyle_PreformattedText());

                    dict.Add(LibraryStyleEnum.Figure, () => GenerateStyle_Figure());
                    dict.Add(LibraryStyleEnum.Table, () => GenerateStyle_Table());
                    dict.Add(LibraryStyleEnum.Drawing, () => GenerateStyle_Drawing());

                    dict.Add(LibraryStyleEnum.Heading, () => GenerateStyle_Heading());

                    dict.Add(LibraryStyleEnum.Heading1, () => GenerateStyle_Heading1(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading2, () => GenerateStyle_Heading2(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading3, () => GenerateStyle_Heading3(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading4, () => GenerateStyle_Heading4(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading5, () => GenerateStyle_Heading5(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading6, () => GenerateStyle_Heading6(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading7, () => GenerateStyle_Heading7(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading8, () => GenerateStyle_Heading8(headingNumberingInstance_none));
                    dict.Add(LibraryStyleEnum.Heading9, () => GenerateStyle_Heading9(headingNumberingInstance_none));

                    dict.Add(LibraryStyleEnum.Index, () => GenerateStyle_Index());
                    dict.Add(LibraryStyleEnum.Contents1, () => GenerateStyle_Contents1());
                    dict.Add(LibraryStyleEnum.IndexHeading, () => GenerateStyle_IndexHeading());
                    dict.Add(LibraryStyleEnum.ContentsHeading, () => GenerateStyle_ContentsHeading());

                    _StyleIdToStyleDict = dict;
                }
                return _StyleIdToStyleDict;
            }
        }

        public Style GenerateStyle_Normal()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal" };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl();
            SuppressAutoHyphens suppressAutoHyphens1 = new SuppressAutoHyphens() { Val = true };
            BiDi biDi1 = new BiDi() { Val = false };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "0", After = "0" };
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            styleParagraphProperties1.Append(widowControl1);
            styleParagraphProperties1.Append(suppressAutoHyphens1);
            styleParagraphProperties1.Append(biDi1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(justification1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif", EastAsia = "Noto Serif CJK SC", ComplexScript = "Lohit Devanagari" };
            Color color1 = new Color() { Val = "auto" };
            Kern kern1 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { Val = "en-GB", EastAsia = "zh-CN", Bidi = "hi-IN" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(kern1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);
            styleRunProperties1.Append(languages1);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading" };
            StyleName styleName1 = new StyleName() { Val = "Heading" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            KeepNext keepNext1 = new KeepNext() { Val = true };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "120" };

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Liberation Sans", HighAnsi = "Liberation Sans", EastAsia = "Noto Sans CJK SC", ComplexScript = "Lohit Devanagari" };
            FontSize fontSize1 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading1(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };
            StyleName styleName1 = new StyleName() { Val = "Heading 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "240", After = "120" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "36" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "36" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading2(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2" };
            StyleName styleName1 = new StyleName() { Val = "Heading 2" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "200", After = "120" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 1 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading3(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading3" };
            StyleName styleName1 = new StyleName() { Val = "Heading 3" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 2 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "140", After = "120" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 2 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading4(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading4" };
            StyleName styleName1 = new StyleName() { Val = "Heading 4" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 3 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "120", After = "120" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 3 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(italic1);
            styleRunProperties1.Append(italicComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading5(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading5" };
            StyleName styleName1 = new StyleName() { Val = "Heading 5" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 4 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "120", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 4 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading6(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading6" };
            StyleName styleName1 = new StyleName() { Val = "Heading 6" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 5 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "60", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 5 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(italic1);
            styleRunProperties1.Append(italicComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading7(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading7" };
            StyleName styleName1 = new StyleName() { Val = "Heading 7" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 6 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "60", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 6 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading8(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading8" };
            StyleName styleName1 = new StyleName() { Val = "Heading 8" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 7 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "60", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 7 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(italic1);
            styleRunProperties1.Append(italicComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Heading9(NumberingInstance numberingInstance)
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading9" };
            StyleName styleName1 = new StyleName() { Val = "Heading 9" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 8 };
            NumberingId numberingId1 = new NumberingId() { Val = numberingInstance.NumberID };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "60", After = "60" };
            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 8 };

            styleParagraphProperties1.Append(numberingProperties1);
            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_InternetLink()
        {
            Style style1 = new Style() { Type = StyleValues.Character, StyleId = "InternetLink" };
            StyleName styleName1 = new StyleName() { Val = "Hyperlink" };

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Color color1 = new Color() { Val = "000080" };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };
            Languages languages1 = new Languages() { Val = "zxx", EastAsia = "zxx", Bidi = "zxx" };

            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(underline1);
            styleRunProperties1.Append(languages1);

            style1.Append(styleName1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Title()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Title" };
            StyleName styleName1 = new StyleName() { Val = "Title" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            styleParagraphProperties1.Append(justification1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "56" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "56" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Subtitle()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Subtitle" };
            StyleName styleName1 = new StyleName() { Val = "Subtitle" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "TextBody" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "60", After = "120" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            styleParagraphProperties1.Append(spacingBetweenLines1);
            styleParagraphProperties1.Append(justification1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize1 = new FontSize() { Val = "36" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "36" };

            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(nextParagraphStyle1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_PreformattedText()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "PreformattedText" };
            StyleName styleName1 = new StyleName() { Val = "Preformatted Text" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "0", After = "0" };

            styleParagraphProperties1.Append(spacingBetweenLines1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Liberation Mono", HighAnsi = "Liberation Mono", EastAsia = "Noto Sans Mono CJK SC", ComplexScript = "Liberation Mono" };
            FontSize fontSize1 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Figure()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Figure" };
            StyleName styleName1 = new StyleName() { Val = "Figure" };
            BasedOn basedOn1 = new BasedOn() { Val = "Caption" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Table()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Table" };
            StyleName styleName1 = new StyleName() { Val = "Table" };
            BasedOn basedOn1 = new BasedOn() { Val = "Caption" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Drawing()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Drawing" };
            StyleName styleName1 = new StyleName() { Val = "Table of Figures" };
            BasedOn basedOn1 = new BasedOn() { Val = "Caption" };
            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Index()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Index" };
            StyleName styleName1 = new StyleName() { Val = "Index" };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers1 = new SuppressLineNumbers();

            styleParagraphProperties1.Append(suppressLineNumbers1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts1 = new RunFonts() { ComplexScript = "Lohit Devanagari" };
            Languages languages1 = new Languages() { Val = "zxx", EastAsia = "zxx", Bidi = "zxx" };

            styleRunProperties1.Append(runFonts1);
            styleRunProperties1.Append(languages1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_Contents1()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Contents1" };
            StyleName styleName1 = new StyleName() { Val = "TOC 1" };
            BasedOn basedOn1 = new BasedOn() { Val = "Index" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Clear, Position = 709 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9638 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            Indentation indentation1 = new Indentation() { Left = "0", Hanging = "0" };

            styleParagraphProperties1.Append(tabs1);
            styleParagraphProperties1.Append(indentation1);
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

            return style1;
        }

        public Style GenerateStyle_IndexHeading()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "IndexHeading" };
            StyleName styleName1 = new StyleName() { Val = "Index Heading" };
            BasedOn basedOn1 = new BasedOn() { Val = "Heading" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers1 = new SuppressLineNumbers();
            Indentation indentation1 = new Indentation() { Left = "0", Hanging = "0" };

            styleParagraphProperties1.Append(suppressLineNumbers1);
            styleParagraphProperties1.Append(indentation1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public Style GenerateStyle_ContentsHeading()
        {
            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "ContentsHeading" };
            StyleName styleName1 = new StyleName() { Val = "TOC Heading" };
            BasedOn basedOn1 = new BasedOn() { Val = "IndexHeading" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SuppressLineNumbers suppressLineNumbers1 = new SuppressLineNumbers();
            Indentation indentation1 = new Indentation() { Left = "0", Hanging = "0" };

            styleParagraphProperties1.Append(suppressLineNumbers1);
            styleParagraphProperties1.Append(indentation1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(boldComplexScript1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(basedOn1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);
            return style1;
        }

        public SdtBlock GenerateSdtBlock(string title = "Table of Contents")
        {
            SdtBlock sdtBlock1 = new SdtBlock();

            SdtProperties sdtProperties1 = new SdtProperties();

            SdtContentDocPartObject sdtContentDocPartObject1 = new SdtContentDocPartObject();
            DocPartGallery docPartGallery1 = new DocPartGallery() { Val = "Table of Contents" };
            DocPartUnique docPartUnique1 = new DocPartUnique() { Val = true };

            sdtContentDocPartObject1.Append(docPartGallery1);
            sdtContentDocPartObject1.Append(docPartUnique1);

            sdtProperties1.Append(sdtContentDocPartObject1);

            SdtContentBlock sdtContentBlock1 = new SdtContentBlock();

            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "ContentsHeading" };
            SuppressLineNumbers suppressLineNumbers1 = new SuppressLineNumbers();
            Indentation indentation1 = new Indentation() { Left = "0", Hanging = "0" };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Bold bold1 = new Bold();
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            paragraphMarkRunProperties1.Append(bold1);
            paragraphMarkRunProperties1.Append(bold2);
            paragraphMarkRunProperties1.Append(boldComplexScript1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(suppressLineNumbers1);
            paragraphProperties1.Append(indentation1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize2 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

            runProperties1.Append(bold3);
            runProperties1.Append(boldComplexScript2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = title;

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Paragraph paragraph2 = new Paragraph();

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Contents1" };
            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run2.Append(fieldChar1);

            Run run3 = new Run();

            RunProperties runProperties2 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "IndexLink" };

            runProperties2.Append(runStyle1);
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " TOC \\f \\o \"1-9\" \\h";

            run3.Append(runProperties2);
            run3.Append(fieldCode1);

            Run run4 = new Run();

            RunProperties runProperties3 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "IndexLink" };

            runProperties3.Append(runStyle2);
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run4.Append(runProperties3);
            run4.Append(fieldChar2);

            Hyperlink hyperlink1 = new Hyperlink() { Anchor = "__RefHeading___Toc552_1129335736" };

            Run run5 = new Run();

            RunProperties runProperties4 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "IndexLink" };

            runProperties4.Append(runStyle3);
            Text text2 = new Text();
            text2.Text = "h1";
            TabChar tabChar1 = new TabChar();
            Text text3 = new Text();
            text3.Text = "1";

            run5.Append(runProperties4);
            run5.Append(text2);
            run5.Append(tabChar1);
            run5.Append(text3);

            hyperlink1.Append(run5);

            Run run6 = new Run();

            RunProperties runProperties5 = new RunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "IndexLink" };

            runProperties5.Append(runStyle4);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run6.Append(runProperties5);
            run6.Append(fieldChar3);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);
            paragraph2.Append(run3);
            paragraph2.Append(run4);
            paragraph2.Append(hyperlink1);
            paragraph2.Append(run6);

            sdtContentBlock1.Append(paragraph1);
            sdtContentBlock1.Append(paragraph2);

            sdtBlock1.Append(sdtProperties1);
            sdtBlock1.Append(sdtContentBlock1);
            return sdtBlock1;
        }

    }

}