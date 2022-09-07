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
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Runtime.CompilerServices;

namespace SearchAThing.DocX
{


    public static partial class DocXExt
    {

        /// <summary>
        /// retrieve existing or create a new main document part with its document and body
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static MainDocumentPart GetMainDocumentPart(this WordprocessingDocument doc)
        {
            if (doc.MainDocumentPart is null)
            {
                doc.AddMainDocumentPart();

                doc.MainDocumentPart!.Document = new Document { Body = new Body() };
            }

            return doc.MainDocumentPart!;
        }

        /// <summary>
        /// retrieve document associated with doc of the main document part
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static Document GetDocument(this WordprocessingDocument doc) => doc.GetMainDocumentPart().Document;

        /// <summary>
        /// retrieve body associated with doc document of the main document part
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static Body GetBody(this WordprocessingDocument doc) => doc.GetDocument().Body!;

        /// <summary>
        /// retrieve or create style definitions parts of the doc main document part
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static StyleDefinitionsPart GetStyleDefinitionsPart(this WordprocessingDocument doc)
        {
            var mainDocumentPart = doc.GetMainDocumentPart();

            if (mainDocumentPart.StyleDefinitionsPart is null)
            {
                var styleDefinitionPart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                styleDefinitionPart.Styles = new Styles();
            }

            return mainDocumentPart.StyleDefinitionsPart!;
        }

        /// <summary>
        /// retrieve or create numbering definitions parts of the doc main document part
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static NumberingDefinitionsPart GetNumberingDefinitionsPart(this WordprocessingDocument doc)
        {
            var mainDocumentPart = doc.GetMainDocumentPart();

            if (mainDocumentPart.NumberingDefinitionsPart is null)
            {
                var numberingDefinitionsPart = mainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                numberingDefinitionsPart.Numbering = new Numbering();
            }

            return mainDocumentPart.NumberingDefinitionsPart!;
        }

        /// <summary>
        /// retrieve or create document settings parts of the doc main document part
        /// </summary>
        public static DocumentSettingsPart GetDocumentSettingsPart(this WordprocessingDocument doc)
        {
            var mainDocumentPart = doc.GetMainDocumentPart();

            if (mainDocumentPart.DocumentSettingsPart is null)
            {
                var documentSettingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>();
                documentSettingsPart.Settings = new Settings();
            }

            return mainDocumentPart.DocumentSettingsPart!;
        }

        
        public static WordprocessingDocument SetDocDefaults(this WordprocessingDocument doc,
            string? runFontName = null,
            System.Drawing.Color? runFontColor = null,
            double? fontSizePt = null,
            SpacingBetweenLinesOptions? spacingBetweenLinesOpts = null,
            IndentationOptions? indentationOpts = null,
            JustificationValues? justification = null)
        {
            var docDefaults = doc.GetDocDefaults(createIfNotExists: true)!;

            if (runFontName is not null || runFontColor is not null || fontSizePt is not null)
            {
                var runPropertiesBaseStyle = docDefaults
                    .GetRunPropertiesDefault(createIfNotExists: true)!
                    .GetRunPropertiesBaseStyle(createIfNotExists: true)!;

                if (runFontName is not null)
                    runPropertiesBaseStyle.GetRunFonts(createIfNotExists: true)!.Ascii = runFontName;

                if (runFontColor is not null)
                    runPropertiesBaseStyle.GetRunColor(createIfNotExists: true)!.Val = runFontColor.ToWPColorString();

                if (fontSizePt is not null)
                    runPropertiesBaseStyle.GetFontSize(createIfNotExists: true)!.Val =
                        fontSizePt.Value.PtToHalfPoint().ToString(CultureInfo.InvariantCulture);
            }

            if (spacingBetweenLinesOpts is not null)
            {
                var spacingBetweenLines = docDefaults
                    .GetParagraphPropertiesDefault(createIfNotExists: true)!
                    .GetParagraphPropertiesBaseStyle(createIfNotExists: true)!
                    .GetSpacingBetweenLines(createIfNotExists: true)!;

                spacingBetweenLines.ApplyOpts(spacingBetweenLinesOpts);
            }

            if (indentationOpts is not null)
            {
                var indentation = docDefaults
                    .GetParagraphPropertiesDefault(createIfNotExists: true)!
                    .GetParagraphPropertiesBaseStyle(createIfNotExists: true)!
                    .GetIndentation(createIfNotExists: true)!;

                indentation.ApplyOpts(indentationOpts);
            }

            if (justification is not null)
            {
                var _justification = docDefaults
                    .GetParagraphPropertiesDefault(createIfNotExists: true)!
                    .GetParagraphPropertiesBaseStyle(createIfNotExists: true)!
                    .GetJustification(createIfNotExists: true)!;

                _justification.Val = justification.Value;
            }

            return doc;
        }

        public static DocDefaults? GetDocDefaults(this WordprocessingDocument doc, bool createIfNotExists = false)
        {
            var styles = doc.GetStyles(createIfNotExists: true)!;

            var docDefaults = styles.Elements<DocDefaults>().FirstOrDefault();
            if (docDefaults is null && createIfNotExists)
            {
                docDefaults = new DocDefaults();
                styles.Append(docDefaults);
            }

            return styles.DocDefaults;
        }


        /// <summary>
        /// insert element into body before main sectionproperties
        /// </summary>
        /// <param name="body">document body</param>
        /// <param name="element">element to insert before man section properties</param>
        /// <param name="doc">(optional) doc ref</param>        
        /// <returns>element</returns>
        public static T AppendBeforeMainSection<T>(this Body body, T element, WordprocessingDocument? doc = null) where T : OpenXmlElement
        {
            if (doc is null) doc = element.GetWordprocessingDocument();

            var lastElement = body.LastChild;
            if (lastElement is not null && lastElement.GetType() == typeof(SectionProperties))
            {
                body.InsertAt(element, body.ChildElements.Count - 1);
            }
            else
            {
                var mainSectionProperties = (SectionProperties)doc.GetMainSectionProperties(createIfNotExists: true, body)!;
                var mainSectionPropertiesIdx = (int)mainSectionProperties.GetIndex()!;
                body.InsertAt(element, mainSectionPropertiesIdx);
            }

            return element;
        }

        public static string DocumentOuterXML(this WordprocessingDocument doc) => doc.GetDocument().OuterXml;        

        /// <summary>
        /// retrieve max id of document properties
        /// </summary>
        public static uint GetMaxDocPrId(this WordprocessingDocument doc)
        {
            var q = doc
                .MainDocumentPart?
                .RootElement?
                .Descendants<DocProperties>()
                .Select(x => x.Id);

            if (q.Any())
                return q.Max(x => x is null ? 0u : x.Value);

            return 0;
        }

        /// <summary>
        /// release Library related resources
        /// </summary>
        /// <param name="doc">wodprocessing document</param>
        public static void Finalize(this WordprocessingDocument doc)
        {
            WrapperManager.Release(doc);
        }

        /// <summary>
        /// retrieve wrapper reference associated with this document.
        /// the wrapper contains some dynamic reference that need to be release with doc.Finalize()
        /// </summary>
        /// <param name="doc">wordprocessing document</param>
        /// <returns>wrapper reference associated with this document</returns>
        public static WrapperManager GetWrapperRef(this WordprocessingDocument doc) =>
            WrapperManager.GetWrapperRef(doc);

        //==================================================================
        //
        // FORWARDERS
        //
        //==================================================================

        /// <summary>
        /// add a break after last body element
        /// </summary>
        /// <param name="doc">wordprocessing doc</param>
        /// <param name="type">type of break</param>
        public static void AddBreak(this WordprocessingDocument doc, BreakValues type = BreakValues.Page) =>
            doc.GetLastElement()?.AddBreak(type);

        /// <summary>
        /// add table at end of body
        /// </summary>
        /// <param name="doc">wordprocessing document</param>
        /// <param name="tableWidthPercent">table width percent (0..100)</param>
        /// <param name="align">table alignment</param>
        /// <returns>table</returns>
        public static Table AddTable(this WordprocessingDocument doc,
            double? tableWidthPercent = null,
            TableRowAlignmentValues align = TableRowAlignmentValues.Left) =>
            doc.GetLastParagraph(createIfNotExists: true)!.AddTable(tableWidthPercent, align);



    }

    public static partial class DocXToolkit
    {

        /// <summary>
        /// note: use doc.Finalize() when finished to release Library resources
        /// </summary>
        /// <param name="docPathfilename">pathfilename which save document</param>
        /// <returns>wordprocessing document</returns>
        public static WordprocessingDocument Create(string docPathfilename)
        {
            var doc = WordprocessingDocument.Create(docPathfilename, WordprocessingDocumentType.Document);

            var mainDocumentPart = doc.GetMainDocumentPart();

            var styleDefinitionPart = doc.GetStyleDefinitionsPart();

            var numberingDefinitionPart = doc.GetNumberingDefinitionsPart();

            var sectionProperties = doc.GetMainSectionProperties(createIfNotExists: true);

            var documentSettingsPart = doc.GetDocumentSettingsPart();

            var docDefaults = doc.GetDocDefaults();

            #region package properties
            doc.PackageProperties.Creator = "username";
            doc.PackageProperties.Title = "";
            doc.PackageProperties.Subject = "";
            doc.PackageProperties.Keywords = "";
            doc.PackageProperties.Description = "";
            doc.PackageProperties.Revision = "1";
            doc.PackageProperties.Created = DateTime.Now;
            doc.PackageProperties.Modified = DateTime.Now;
            doc.PackageProperties.LastModifiedBy = "username";
            #endregion         

            return doc;
        }

        /// <summary>
        /// note: use doc.Finalize() when finished to release Library resources
        /// </summary>
        /// <param name="docPathfilename">pathfilename which save document</param>
        /// <param name="isEditable">if false open in readonly</param>
        /// <returns>wordprocessing document</returns>
        public static WordprocessingDocument Open(string docPathfilename, bool isEditable = true) =>
            WordprocessingDocument.Open(docPathfilename, isEditable);

    }

}