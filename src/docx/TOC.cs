namespace SearchAThing.DocX;

public static partial class DocXExt
{

    /// <summary>
    /// add toc ad end of body or after given paragraphBefore
    /// </summary>
    /// <param name="doc">word processing document</param>
    /// <param name="tocTitle">title of toc</param>
    /// <param name="paragraphBefore">(optional) if not null toc will placed just after given paragraphBefore</param>
    /// <param name="updateOnOpen">(optional) if true, as default, toc marked dirty to allow rebuild dialog on open</param>
    /// <returns>toc</returns>
    public static SdtBlock AddToc(this WordprocessingDocument doc,
        string tocTitle = "Table of Contents",
        Paragraph? paragraphBefore = null,
        bool updateOnOpen = true)
    {
        var lib = doc.GetWrapperRef();

        var toc = lib.GenerateSdtBlock(tocTitle);

        var body = doc.GetBody();

        if (paragraphBefore is not null)
            body.InsertAfter(toc, paragraphBefore);
        else
            body.AppendBeforeMainSection(toc, doc);

        lib.IntegrateRequiredStyles(toc);

        if (updateOnOpen) doc.SetUpdateTOCOnOpen();

        return toc;
    }

    /// <summary>
    /// add toc after given paragraph
    /// </summary>
    /// <param name="paragraph">toc will placed just after given paragraph</param>
    /// <param name="tocTitle">title of toc</param>        
    /// <param name="updateOnOpen">(optional) if true, as default, toc marked dirty to allow rebuild dialog on open</param>
    /// <returns>toc</returns>
    public static SdtBlock AddToc(this Paragraph paragraph,
        string tocTitle = "Table of Contents",
        bool updateOnOpen = true)
    {
        var doc = paragraph.GetWordprocessingDocument();

        return doc.AddToc(tocTitle, paragraph, updateOnOpen);
    }

    internal static WordprocessingDocument SetUpdateTOCOnOpen(this WordprocessingDocument doc)
    {
        var settings = doc.GetDocumentSettingsPart().Settings;

        var updateFieldsOnOpen = new UpdateFieldsOnOpen { Val = true };
        settings.PrependChild(updateFieldsOnOpen);

        return doc;
    }


}