namespace SearchAThing.DocX;

public static partial class DocXExt
{

    internal static RunPropertiesDefault? GetRunPropertiesDefault(this DocDefaults docDefaults,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        docDefaults.GetOrCreate<RunPropertiesDefault>(createIfNotExists, insertAtIdx);

    internal static RunPropertiesBaseStyle? GetRunPropertiesBaseStyle(this RunPropertiesDefault runPropertiesDefault,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        runPropertiesDefault.GetOrCreate<RunPropertiesBaseStyle>(createIfNotExists, insertAtIdx);

    internal static RunFonts? GetRunFonts(this RunPropertiesBaseStyle runPropertiesBaseStyle,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        runPropertiesBaseStyle.GetOrCreate<RunFonts>(createIfNotExists, insertAtIdx);

    internal static Color? GetRunColor(this RunPropertiesBaseStyle runPropertiesBaseStyle,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        runPropertiesBaseStyle.GetOrCreate<Color>(createIfNotExists, insertAtIdx);

    internal static FontSize? GetFontSize(this RunPropertiesBaseStyle runPropertiesBaseStyle,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        runPropertiesBaseStyle.GetOrCreate<FontSize>(createIfNotExists, insertAtIdx);

    internal static ParagraphPropertiesDefault? GetParagraphPropertiesDefault(this DocDefaults docDefaults,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        docDefaults.GetOrCreate<ParagraphPropertiesDefault>(createIfNotExists, insertAtIdx);

    internal static ParagraphPropertiesBaseStyle? GetParagraphPropertiesBaseStyle(this ParagraphPropertiesDefault paragraphPropertiesDefault,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphPropertiesDefault.GetOrCreate<ParagraphPropertiesBaseStyle>(createIfNotExists, insertAtIdx);

    internal static SpacingBetweenLines? GetSpacingBetweenLines(this ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphPropertiesBaseStyle.GetOrCreate<SpacingBetweenLines>(createIfNotExists, insertAtIdx);

    internal static Indentation? GetIndentation(this ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphPropertiesBaseStyle.GetOrCreate<Indentation>(createIfNotExists, insertAtIdx);

    internal static Justification? GetJustification(this ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphPropertiesBaseStyle.GetOrCreate<Justification>(createIfNotExists, insertAtIdx);

    internal static HeaderPart CreateHeaderPart(this WordprocessingDocument doc) => doc
        .GetMainDocumentPart()
        .AddNewPart<HeaderPart>();

    internal static FooterPart CreateFooterPart(this WordprocessingDocument doc) => doc
        .GetMainDocumentPart()
        .AddNewPart<FooterPart>();


    public static Indentation ApplyOpts(this Indentation indentation, IndentationOptions opts)
    {
        if (opts.StartMM is not null)
            indentation.Start = opts.StartMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

        if (opts.EndMM is not null)
            indentation.End = opts.EndMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

        if (opts.HangingMM is not null)
            indentation.Hanging = opts.HangingMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

        return indentation;
    }

    public static SpacingBetweenLines ApplyOpts(this SpacingBetweenLines spacingBetweenLines, SpacingBetweenLinesOptions opts)
    {
        if (opts.AfterMM is not null)
            spacingBetweenLines.After = opts.AfterMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

        if (opts.BeforeMM is not null)
            spacingBetweenLines.Before = opts.BeforeMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

        if (opts.LineHeightMM is not null)
        {
            spacingBetweenLines.Line = opts.LineHeightMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);
            spacingBetweenLines.LineRule = LineSpacingRuleValues.Exact;
        }

        return spacingBetweenLines;
    }

    /// <summary>
    /// retrieve Styles associated with style definitions parts of the doc main document part
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static Styles? GetStyles(this WordprocessingDocument doc, bool createIfNotExists = false)
    {
        var styleDefinitionsPart = GetStyleDefinitionsPart(doc);

        if (styleDefinitionsPart.Styles is null && createIfNotExists) styleDefinitionsPart.Styles = new Styles();

        return styleDefinitionsPart.Styles;
    }

    /// <summary>
    /// copy section properties from given template to target
    /// </summary>
    /// <param name="sectionProperties">target section property</param>
    /// <param name="templateSectionProperties">template section property</param>
    /// <param name="overwriteExistingMembers">if true template member will overwrite existing target members;
    /// if false (default) existing target member setting will maintained and added those missing</param>
    /// <returns>target section property with modified parts</returns>
    internal static SectionProperties CopyFrom(this SectionProperties sectionProperties,
        SectionProperties templateSectionProperties,
        bool overwriteExistingMembers = false)
    {
        var pageMargin = sectionProperties.GetOrCreate<PageMargin>(createIfNotExists: false);
        if (pageMargin is null || overwriteExistingMembers)
        {
            var templatePageMargin = templateSectionProperties.GetOrCreate<PageMargin>(createIfNotExists: false);

            if (templatePageMargin is not null)
            {
                if (pageMargin is not null) pageMargin.Remove();

                var templatePageMarginClone = (PageMargin)templatePageMargin.Clone();

                sectionProperties.Append(templatePageMarginClone);
            }
        }

        var pageSize = sectionProperties.GetOrCreate<PageSize>(createIfNotExists: false);
        if (pageSize is null || overwriteExistingMembers)
        {
            var templatePageSize = templateSectionProperties.GetOrCreate<PageSize>(createIfNotExists: false);

            if (templatePageSize is not null)
            {
                if (pageSize is not null) pageSize.Remove();

                var templatePageSizeClone = (PageSize)templatePageSize.Clone();

                sectionProperties.Append(templatePageSizeClone);
            }
        }

        var headerRef = sectionProperties.GetOrCreate<HeaderReference>(createIfNotExists: false);
        if (headerRef is null || overwriteExistingMembers)
        {
            var templateHeaderRef = templateSectionProperties.GetOrCreate<HeaderReference>(createIfNotExists: false);

            if (templateHeaderRef is not null)
            {
                if (headerRef is not null) headerRef.Remove();

                var templateHeaderRefClone = (HeaderReference)templateHeaderRef.Clone();

                sectionProperties.InsertAt(templateHeaderRefClone, 0);
            }
        }

        var footerRef = sectionProperties.GetOrCreate<FooterReference>(createIfNotExists: false);
        if (footerRef is null || overwriteExistingMembers)
        {
            var templateFooterRef = templateSectionProperties.GetOrCreate<FooterReference>(createIfNotExists: false);

            if (templateFooterRef is not null)
            {
                if (footerRef is not null) footerRef.Remove();

                var templateFooterRefClone = (FooterReference)templateFooterRef.Clone();

                sectionProperties.InsertAt(templateFooterRefClone, 0);
            }
        }

        return sectionProperties;
    }


    public static ParagraphStyleId? GetStyleId(this ParagraphProperties paragraphProperties,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphProperties.GetOrCreate<ParagraphStyleId>(createIfNotExists, insertAtIdx);

    public static string? GetStyleName(this Paragraph paragraph) => paragraph
        .GetProperties()?
        .GetStyleName();

    public static string? GetStyleName(this ParagraphProperties paragraphProperties) => paragraphProperties
        .GetFirstChild<ParagraphStyleId>()?
        .Val;

    public static Style? GetStyle(this Paragraph paragraph) =>
        paragraph
        .GetWordprocessingDocument()
        .GetStyleByName(paragraph.GetStyleName());

    /// <summary>
    /// retrieve next paragraph style from this one
    /// </summary>
    /// <param name="paragraph">paragraph for which retrieve next paragraph style</param>
    /// <param name="doc">(optional) if null WordprocessingDocument will retrieved by parent</param>
    /// <returns>Style of next paragraph</returns>
    public static Style? GetNextParagraphStyle(this Paragraph paragraph, WordprocessingDocument? doc = null)
    {
        if (doc is null) doc = paragraph.GetWordprocessingDocument();

        return doc.GetStyleByName(paragraph.GetStyle()?.GetNextParagraphStyleName());
    }

    public static StyleRunProperties? GetStyleRunProperties(this ParagraphProperties paragraphProperties,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphProperties.GetOrCreate<StyleRunProperties>(createIfNotExists, insertAtIdx);

    public static SectionProperties? GetSectionProperties(this ParagraphProperties paragraphProperties,
        bool createIfNotExists = false,
        int? insertAtIdx = null) =>
        paragraphProperties.GetOrCreate<SectionProperties>(createIfNotExists, insertAtIdx);

    public static IEnumerable<Style> EnumStyles(this WordprocessingDocument doc) =>
        doc.GetStyleDefinitionsPart().Styles.OfType<Style>();

    public static IEnumerable<(Style style, string id)> GetStyleWithIds(this WordprocessingDocument doc,
        StyleValues? typeOfStyle = null) =>
        doc
        .EnumStyles()
        .Where(style =>
            (typeOfStyle is null || (style.Type != null && style.Type.Value == typeOfStyle)) &&
            style.StyleId != null && style.StyleId.Value != null)
        .Select(style => (style, id: style.StyleId!.Value!));

    public static string? GetStyleName(this Style style) => style.StyleName?.Val;

    public static IEnumerable<(Style style, string name)> GetStyleWithNames(this WordprocessingDocument doc,
        StyleValues? typeOfStyle = null) =>
        doc
        .EnumStyles()
        .Where(style => typeOfStyle is null || (style.Type != null && style.Type.Value == typeOfStyle))
        .Select(style => (style, name: style.GetStyleName()!));

    //

    /// <summary>
    /// states if document contains given syle id
    /// </summary>
    /// <param name="doc">wordprocessing document</param>
    /// <param name="styleId">style id</param>
    /// <returns>true if style id present in the document styles</returns>
    public static bool HasStyleId(this WordprocessingDocument doc, string styleId) =>
        doc.GetStyleWithIds().Any(nfo => nfo.id == styleId);

    /// <summary>
    /// retrieve style by given style id
    /// </summary>
    /// <param name="doc">wordprocessing doc</param>
    /// <param name="styleId">style id</param>
    /// <param name="integrateIfNotFound">if true (default) searches for predefined styles</param>
    /// <returns>style</returns>
    public static Style? GetStyleById(this WordprocessingDocument doc,
        string styleId,
        bool integrateIfNotFound = true)
    {
        var style = doc.GetStyleWithIds().FirstOrDefault(w => w.id == styleId).style;

        if (integrateIfNotFound && style is null)
        {
            var lib = WrapperManager.GetWrapperRef(doc);
            var qLib = lib.TryGetStyle(styleId);
            if (qLib is not null)
            {
                style = qLib;
                var styles = doc.GetStyles(createIfNotExists: true)!;
                styles.Append(style);
            }
        }

        return style;
    }

    /// <summary>
    /// retrieve on of available style from library
    /// </summary>
    /// <param name="doc">wordprocessing document</param>
    /// <param name="style">style type</param>
    /// <returns>style</returns>
    public static Style GetPredefinedStyle(this WordprocessingDocument doc, LibraryStyleEnum style) =>
        doc.GetStyleById(style.ToString())!;

    public static Style? GetStyleByName(this WordprocessingDocument doc, string? styleName) =>
        doc.GetStyleWithNames().FirstOrDefault(w => w.name == styleName).style;

    //                

    /// <summary>
    /// Adds paragraph style
    /// </summary>
    /// <param name="doc">Wordprocessing wrapper</param>
    /// <param name="styleName">paragraph style name</param>
    /// <param name="runFontName">RUN font name</param>
    /// <param name="runFontColor">RUN font color (if null "auto" will used)</param>
    /// <param name="runFontSizePt">RUN font size in pt</param>
    /// <param name="spacingBetweenLinesOpts">paragrah line spacing options</param>
    /// <param name="indentationOpts">paragrah indentation options</param>
    /// <param name="justification">paragrah justification</param>
    /// <param name="styleId">paragraph style id ( if not specified name will given )</param>        
    /// <param name="basedOn">style which this is based on</param>        
    /// <returns>new paragraph style</returns>
    public static Style AddParagraphSyle(this WordprocessingDocument doc,
        string styleName,
        string? runFontName = null,
        System.Drawing.Color? runFontColor = null,
        double? runFontSizePt = null,
        SpacingBetweenLinesOptions? spacingBetweenLinesOpts = null,
        IndentationOptions? indentationOpts = null,
        JustificationValues? justification = null,
        Style? basedOn = null,
        string? styleId = null)
    {
        Style style;

        if (basedOn is not null)
        {
            style = (Style)basedOn.Clone();
        }
        else style = new Style
        {
            Type = StyleValues.Paragraph,
        };

        style.StyleId = styleId == null ? styleName : styleId;

        {
            var _styleName = new StyleName { Val = styleName };
            var primaryStyle = new PrimaryStyle();

            style.Append(_styleName);
            style.Append(primaryStyle);
        }

        if (basedOn is not null)
        {
            var basedOnObj = new BasedOn { Val = basedOn.StyleName?.Val };
            style.Append(basedOnObj);
        }

        {
            var styleParagraphProperties = style.GetOrCreate<StyleParagraphProperties>(createIfNotExists: true)!;

            if (spacingBetweenLinesOpts is not null)
            {
                var spacingBetweenLines = styleParagraphProperties.GetOrCreate<SpacingBetweenLines>(createIfNotExists: true)!;

                spacingBetweenLines.ApplyOpts(spacingBetweenLinesOpts);
            }

            if (indentationOpts is not null)
            {
                var indentation = styleParagraphProperties.GetOrCreate<Indentation>(createIfNotExists: true)!;

                indentation.ApplyOpts(indentationOpts);
            }

            if (justification is not null)
            {
                styleParagraphProperties.Justification = new Justification
                {
                    Val = justification
                };
            }
        }

        {
            var styleRunProperties = style.GetOrCreate<StyleRunProperties>(createIfNotExists: true)!;

            {
                var runFonts = styleRunProperties.GetOrCreate<RunFonts>(createIfNotExists: true)!;
                if (runFontName is not null)
                    runFonts.Ascii = runFontName;

                var color = styleRunProperties.GetOrCreate<Color>(createIfNotExists: true)!;
                if (runFontColor is not null)
                    color.Val = runFontColor.ToWPColorString();

                var fontSize = styleRunProperties.GetOrCreate<FontSize>(createIfNotExists: true)!;
                if (runFontSizePt is not null)
                    fontSize.Val = runFontSizePt.Value.PtToHalfPoint().ToString(CultureInfo.InvariantCulture);
            }
        }

        doc.GetStyleDefinitionsPart().Styles!.Append(style);

        return style;
    }


}