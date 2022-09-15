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

namespace SearchAThing.DocX
{

    public static partial class DocXExt
    {

        /// <summary>
        /// retrieve all document paragraphs
        /// </summary>
        public static IEnumerable<Paragraph> GetParagraphs(this WordprocessingDocument doc) =>
            doc.GetBody().Elements<Paragraph>();

        /// <summary>
        /// retrieve all document paragraphs runs
        /// </summary>        
        public static IEnumerable<Run> GetRuns(this WordprocessingDocument doc) => doc
            .GetParagraphs()
            .SelectMany(paragraph => paragraph.GetRuns());

        /// <summary>
        /// retrieve document last paragraph
        /// </summary>
        public static Paragraph? GetLastParagraph(this WordprocessingDocument doc, bool createIfNotExists = false)
        {
            var body = doc.GetBody();

            var element = body.LastChild;
            if (element is not null)
            {
                for (int idx = body.ChildElements.Count; idx >= 0; --idx)
                {
                    if (element is Paragraph paragraph) return paragraph;

                    element = element?.PreviousSibling<Paragraph>();
                }
            }

            if (createIfNotExists)
                return doc.AddParagraph();
            else
                return null;
        }

        /// <summary>
        /// retrieve next paragraphs of this
        /// </summary>
        /// <param name="paragraph">paragraph for which get the next</param>
        /// <param name="condition">if not null allow to search next paragraph iff meet given condition</param>
        /// <returns>next paragraphs</returns>
        public static IEnumerable<Paragraph> GetNextParagraphs(this Paragraph paragraph, Func<Paragraph, bool>? condition = null)
        {
            var nextParagraph = paragraph.NextSibling<Paragraph>();

            if (nextParagraph is null) yield break;

            if (condition is not null)
            {
                if (condition(nextParagraph)) yield return nextParagraph;

                foreach (var x in nextParagraph.GetNextParagraphs(condition)) yield return x;
            }
            else
                yield return nextParagraph;
        }

        public static string? GetNextParagraphStyleName(this Style style) => style
            .Elements<NextParagraphStyle>()
            .FirstOrDefault()?
            .Val;

        public static ParagraphProperties? GetProperties(this Paragraph paragraph,
            bool createIfNotExists = false,
            int insertAtIdx = 0) =>
            paragraph.GetOrCreate<ParagraphProperties>(createIfNotExists, insertAtIdx);

        /// <summary>
        /// retrieve runs belonging this paragraph
        /// </summary>
        /// <param name="paragraph">paragraph for which retrieve runs</param>
        /// <returns>runs belonging this paragraph</returns>
        public static IEnumerable<Run> GetRuns(this Paragraph paragraph) => paragraph.Elements<Run>();

        public static RunProperties? GetProperties(this Run run,
            bool createIfNotExists = false,
            int insertAtIdx = 0) =>
            run.GetOrCreate<RunProperties>(createIfNotExists, insertAtIdx);

        /// <summary>
        /// set run runProperties
        /// </summary>
        /// <param name="run">run which set run properties</param>
        /// <param name="runProperties">run properties to set to run</param>
        /// <returns>run</returns>
        public static Run SetProperties(this Run run, RunProperties? runProperties = null)
        {
            var qExistingRunProperties = run.GetProperties();
            if (qExistingRunProperties is not null)
                qExistingRunProperties.Remove();

            if (runProperties is not null)
                run.InsertAt(runProperties, 0);

            return run;
        }

        /// <summary>
        /// copy (cloned) this run properties to given dst run
        /// </summary>
        /// <param name="run">this run</param>
        /// <param name="dstRun">destination run</param>
        /// <returns>this run</returns>
        public static Run CopyPropertiesTo(this Run run, Run dstRun) =>
            dstRun.SetProperties((RunProperties?)run.GetProperties()?.Clone());




        /// <summary>
        /// add a new paragraph after given paragraphBefore or to end of document
        /// </summary>
        /// <param name="doc">word processing document</param>
        /// <param name="txt">paragraph initial text</param>
        /// <param name="paragraphBefore">(optional) paragraph before the new one</param>
        /// <param name="elementBefore">(optional) element before the new one</param>
        /// <param name="parent">(optional) specify parent element</param>
        /// <param name="style">(optional) style to apply to new paragraph or if exists a previous one it will inherithed</param>
        /// <param name="action">(optional) execute action on run created with this paragrah if any</param>
        /// <returns>new paragraph</returns>
        public static Paragraph AddParagraph(this WordprocessingDocument doc,
            string? txt = null,
            Style? style = null,
            Paragraph? paragraphBefore = null,
            OpenXmlElement? elementBefore = null,
            OpenXmlElement? parent = null,
            Action<Run>? action = null)
        {
            var paragraph = new Paragraph();

            {
                var body = doc.GetBody();

                var paragraphProperties = paragraph.GetProperties(createIfNotExists: true, insertAtIdx: 0)!;
                string? styleId = null;

                //if (elementBefore is null && paragraphBefore is null) paragraphBefore = doc.GetLastElement();//.GetLastParagraph();

                if (style == null)
                    styleId = paragraphBefore?.GetNextParagraphStyle(doc)?.StyleId?.Value;
                else
                    styleId = style.StyleId;

                if (styleId != null)
                {
                    var paragrahStyleId = paragraphProperties.GetStyleId(createIfNotExists: true);
                    paragrahStyleId!.Val = styleId;


                }

                if (paragraphBefore is not null) elementBefore = paragraphBefore;

                if (elementBefore is null)
                {
                    body.AppendBeforeMainSection(paragraph, doc);
                }

                else
                {
                    var nextNonParagraph = elementBefore
                        .NextSibling(condition: element =>
                            element.GetType() != typeof(Paragraph) &&
                            element.GetType() != typeof(SectionProperties));

                    if (nextNonParagraph is null)
                    {
                        if (parent is not null)
                            parent.Append(paragraph);
                        else
                            elementBefore.InsertAfterSelf(paragraph);
                    }

                    else
                    {
                        var insertIdx = body.ChildElements.Count;

                        body.InsertAt(paragraph, insertIdx);
                    }
                }
            }

            if (action is not null && txt is null) throw new Exception($"action specified but not runs avail");

            if (txt != null) paragraph.AddText(txt, action: action);

            var wrapper = paragraph.GetWrapperRef();
            if (wrapper.numberingEnabled is not null)
                paragraph.SetNumbering(
                    levelIndex: wrapper.numberingLevel,
                    format: wrapper.numberingEnabled.Value);

            return paragraph;
        }

        /// <summary>
        /// set paragraph style
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="style">style to apply</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetParagraphStyle(this Paragraph paragraph, Style style)
        {
            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true)!;
            paragraphProperties.GetStyleId(createIfNotExists: true)!.Val = style.StyleId;

            return paragraph;
        }

        public static Run SetFontName(this Run run, string fontName)
        {
            var runProperties = run.GetProperties(createIfNotExists: true)!;

            runProperties.RunFonts = new RunFonts { Ascii = fontName };

            return run;
        }

        public static Run SetBold(this Run run, bool bold = true)
        {
            var runProperties = run.GetProperties(createIfNotExists: true)!;

            runProperties.Bold = new Bold { Val = OnOffValue.FromBoolean(bold) };

            return run;
        }

        public static Run SetItalic(this Run run, bool italic = true)
        {
            var runProperties = run.GetProperties(createIfNotExists: true)!;

            runProperties.Italic = new Italic { Val = OnOffValue.FromBoolean(italic) };

            return run;
        }

        /// <summary>
        /// set underline type to the given run
        /// </summary>
        /// <param name="run">run which apply underline</param>
        /// <param name="underline">underline type</param>
        /// <returns>run</returns>
        public static Run SetUnderline(this Run run, UnderlineValues underline = UnderlineValues.Single)
        {
            var runProperties = run.GetProperties(createIfNotExists: true)!;

            runProperties.Underline = new Underline { Val = underline };

            return run;
        }

        /// <summary>
        /// set color to the given run
        /// </summary>
        /// <param name="run">run which apply color</param>
        /// <param name="color">color</param>
        /// <returns>run</returns>
        public static Run SetColor(this Run run, System.Drawing.Color? color = null)
        {
            var runProperties = run.GetProperties(createIfNotExists: true)!;

            runProperties.Color = new Color { Val = color.ToWPColorString() };

            return run;
        }

        /// <summary>
        /// set shading over paragraph extensions
        /// </summary>
        /// <param name="paragraph">paragraph which apply shading</param>
        /// <param name="color">shading color</param>
        /// <param name="pattern">shading pattern type</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetShading(this Paragraph paragraph,
            System.Drawing.Color? color = null,
            ShadingPatternValues pattern = ShadingPatternValues.Clear) =>
            paragraph.SetShading<Paragraph, ParagraphProperties>(color, pattern);

        /// <summary>
        /// set shading over run extensions
        /// </summary>
        /// <param name="run">run which apply shading</param>
        /// <param name="color">shading color</param>
        /// <param name="pattern">shading pattern type</param>
        /// <returns>run</returns>
        public static Run SetShading(this Run run, System.Drawing.Color? color = null,
            ShadingPatternValues pattern = ShadingPatternValues.Clear) =>
            run.SetShading<Run, RunProperties>(color, pattern);

        /// <summary>
        /// set paragraph numbering style
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="levelIndex">level of numbering</param>
        /// <param name="format">type of numbering</param>        
        /// <param name="structured">if true previous level info reported in subsequent</param>        
        /// <param name="restartNumbering">if true restart numbering from first number defined by related abstract numbering; this accomplished by allocating a new numbering instance</param>        
        /// <returns>paragraph</returns>        
        public static Paragraph SetNumbering(this Paragraph paragraph,
            int levelIndex = 1,
            NumberFormatValues format = NumberFormatValues.Bullet,
            bool structured = false,
            bool restartNumbering = false)
        {
            var doc = paragraph.GetWordprocessingDocument();

            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true)!;

            var numberingProperties = paragraphProperties.GetOrCreate<NumberingProperties>(createIfNotExists: true)!;

            if (numberingProperties.NumberingId is not null) numberingProperties.NumberingId.Remove();

            var abstractNum = doc.GetAbstractNum(format, structured, restartNumbering);

            var numberingInstance = doc.GetNumberingInstance(abstractNum);
            //

            numberingProperties.Append(new NumberingLevelReference { Val = levelIndex });
            numberingProperties.Append(new NumberingId { Val = numberingInstance.NumberID });

            return paragraph;
        }


        /// <summary>
        /// set font size of given run
        /// </summary>
        /// <param name="run">run</param>
        /// <param name="fontSizePt">font size ( in points )</param>
        /// <returns>run</returns>
        public static Run SetFontSizePt(this Run run, double fontSizePt)
        {
            var runProperties = run.GetProperties(createIfNotExists: true)!;

            runProperties.FontSize = new FontSize { Val = fontSizePt.PtToHalfPoint().ToString(CultureInfo.InvariantCulture) };

            return run;
        }

        /// <summary>
        /// set text of given paragraph replacing existing if any
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="textStr">new text</param>
        /// <param name="action">action to execute on paragraph run child</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetText(this Paragraph paragraph, string textStr, Action<Run>? action = null)
        {
            var toremove = paragraph.Elements<Run>().ToList();
            foreach (var child in toremove) paragraph.RemoveChild(child);

            var newRun = new Run(new Text(textStr).Act(text => text.Space = SpaceProcessingModeValues.Preserve));

            paragraph.AddChild(newRun);

            if (action != null) action(newRun);

            return paragraph;
        }

        /// <summary>
        /// add a new run to given paragraph
        /// </summary>
        /// <param name="txt">run txt</param>
        /// <param name="paragraph">paragraph which adds a new run</param>
        /// <param name="action">action to execute on new run</param>
        /// <param name="runIdx">(optional) index of run insertion</param>
        /// <returns>paragraph</returns>
        public static Paragraph AddRun(this Paragraph paragraph, string txt, Action<Run>? action = null, int? runIdx = null)
        {
            var run = new Run();
            if (runIdx != null)
                paragraph.InsertAt(run, runIdx.Value);
            else
                paragraph.AppendChild(run);

            run.SetText(txt);

            if (action != null) action(run);

            return paragraph;
        }

        /// <summary>
        /// add given runs to this paragraph
        /// </summary>        
        /// <param name="paragraph">paragraph which adds runs</param>        
        /// <param name="run">already allocated run (without parent)</param>                
        /// <returns>paragraph</returns>
        public static Paragraph AddRun(this Paragraph paragraph, Run run)
        {
            paragraph.Append(run);

            return paragraph;
        }

        /// <summary>
        /// add given runs to this paragraph
        /// </summary>        
        /// <param name="paragraph">paragraph which adds runs</param>        
        /// <param name="runs">already allocated run (without parent)</param>        
        /// <returns>paragraph</returns>
        public static Paragraph AddRuns(this Paragraph paragraph, IEnumerable<Run> runs) =>
            paragraph.Act(_ => runs.Foreach(run => paragraph.Append(run)));

        /// <summary>
        /// add a new run to given paragraph
        /// </summary>
        /// <param name="paragraph">paragraph which adds a new run</param>
        /// <param name="action">action to execute on new run</param>
        /// <param name="runIdx">(optional) index of run insertion</param>
        /// <returns>paragraph</returns>
        public static Paragraph AddRun(this Paragraph paragraph, Action<Run>? action = null, int? runIdx = null)
        {
            var run = new Run();
            if (runIdx != null)
                paragraph.InsertAt(run, runIdx.Value);
            else
                paragraph.AppendChild(run);

            if (action != null) action(run);

            return paragraph;
        }

        /// <summary>
        /// add a new run with a single space to given paragraph
        /// </summary>
        /// <param name="paragraph">paragraph which adds a new run</param>
        /// <param name="action">action to execute on new run</param>
        /// <param name="runIdx">(optional) index of run insertion</param>
        /// <returns>paragraph</returns>
        public static Paragraph AddSpace(this Paragraph paragraph, Action<Run>? action = null, int? runIdx = null) =>
            paragraph.AddRun(" ", action, runIdx);

        internal static Run AddInlineDrawing(this Run run, Drawing drawing) => run.Act(run => run.Append(drawing));

        /// <summary>
        /// add a text to the given run
        /// </summary>
        /// <param name="run">run</param>
        /// <param name="textStr">text to add the run</param>
        /// <returns>run</returns>
        public static Run AddText(this Run run, string textStr) =>
            run.Act(run => run.Append(new Text(textStr).Act(text => text.Space = SpaceProcessingModeValues.Preserve)));

        /// <summary>
        /// add a text to the given paragraph creating a corresponding new run
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="textStr">text to add the paragraph</param>
        /// <param name="action">action to apply to the new run containing added text</param>
        /// <param name="runIdx">(optional) index which insert the new run</param>
        /// <returns>paragraph</returns>
        public static Paragraph AddText(this Paragraph paragraph, string textStr, Action<Run>? action = null, int? runIdx = null) =>
            paragraph.AddRun(action: r => { r.AddText(textStr); action?.Invoke(r); }, runIdx: runIdx);

        /// <summary>
        /// set bold to given value over all paragraph runs
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="bold">bold mode</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetBold(this Paragraph paragraph, bool bold = true) =>
            paragraph.Act(paragraph => paragraph.GetRuns().Foreach(run => run.SetBold(bold)));

        /// <summary>
        /// set italic to given value over all paragraph runs
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="italic">italic mode</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetItalic(this Paragraph paragraph, bool italic = true) =>
            paragraph.Act(paragraph => paragraph.GetRuns().Foreach(run => run.SetItalic(italic)));

        /// <summary>
        /// set underline to given value over all paragraph runs
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="type">underline type</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetUnderline(this Paragraph paragraph, UnderlineValues type = UnderlineValues.Single) =>
            paragraph.Act(paragraph => paragraph.GetRuns().Foreach(run => run.SetUnderline(type)));

        /// <summary>
        /// add a standard field to this paragraph
        /// </summary>
        /// <param name="paragraph">paragraph which adds a field</param>
        /// <param name="field">field type</param>
        /// <param name="action">(optional) action to apply at run which will contains the field</param>
        /// <returns>paragraph</returns>
        public static Paragraph AddField(this Paragraph paragraph, FieldEnum field, Action<Run>? action = null) =>
            paragraph.AddField(field.ToString(), action);

        /// <summary>
        /// add a custom field to this paragraph
        /// </summary>
        /// <param name="paragraph">paragraph which adds a field</param>
        /// <param name="field">field text</param>
        /// <param name="action">(optional) action to apply at run which will contains the field</param>
        /// <returns>paragraph</returns>
        public static Paragraph AddField(this Paragraph paragraph, string field, Action<Run>? action = null)
        {
            paragraph.AddRun(action: run => { action?.Invoke(run); run.Append(new FieldChar { FieldCharType = FieldCharValues.Begin }); });
            paragraph.AddRun(action: run => { action?.Invoke(run); run.Append(new FieldCode { Text = field }); });
            // paragraph.AddRun(action: run => run.Append(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            // paragraph.AddText("xxxxx", action);
            paragraph.AddRun(action: run => { action?.Invoke(run); run.Append(new FieldChar { FieldCharType = FieldCharValues.End }); });

            return paragraph;
        }

        /// <summary>
        /// retrieve run text string
        /// </summary>
        /// <param name="run">run</param>
        /// <param name="createIfNotExists">(optional) if true create the run text child if not exists</param>
        /// <returns>text contained in the run</returns>
        public static string? GetTextStr(this Run run, bool createIfNotExists = false) =>
            run.GetText(createIfNotExists)?.Text;

        /// <summary>
        /// retrieve the text object
        /// </summary>
        /// <param name="run">run</param>
        /// <param name="createIfNotExists">(optional) if true create the run text child if not exists</param>
        /// <param name="insertAtIdx">(optional) insertion index for the new text run</param>        
        /// <returns>text object contained in the run</returns>
        public static Text? GetText(this Run run, bool createIfNotExists = false, int? insertAtIdx = null) =>
            run.GetOrCreate<Text>(createIfNotExists, insertAtIdx, onNew: (text) =>
            {
                text.Space = SpaceProcessingModeValues.Preserve;
            });

        /// <summary>
        /// replace run text
        /// </summary>
        /// <param name="run">run</param>
        /// <param name="newtext">new run text</param>
        /// <returns>run</returns>
        public static Run SetText(this Run run, string newtext)
        {
            run.GetText(createIfNotExists: true)!.Text = newtext;

            return run;
        }

        /// <summary>
        /// set run as subscript
        /// </summary>
        /// <param name="run">run which apply subscript</param>
        /// <returns>run</returns>
        public static Run SetSubscript(this Run run) => run.SetVerticalPosition(VerticalPositionValues.Subscript);

        /// <summary>
        /// set run as superscript
        /// </summary>
        /// <param name="run">run which apply superscript</param>
        /// <returns>run</returns>
        public static Run SetSuperscript(this Run run) => run.SetVerticalPosition(VerticalPositionValues.Superscript);

        /// <summary>
        /// set run vertical position
        /// </summary>
        /// <param name="run">run</param>
        /// <param name="pos">pos</param>
        /// <returns>run with changed vertical position</returns>
        public static Run SetVerticalPosition(this Run run, VerticalPositionValues pos)
        {
            var rp = run.GetProperties(createIfNotExists: true)!;

            var vp = rp.GetOrCreate<VerticalTextAlignment>(createIfNotExists: true)!;

            vp.Val = pos;

            return run;
        }

        /// <summary>
        /// set spacing between lines of paragraph
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="beforeMM">space before (mm)</param>
        /// <param name="afterMM">space after (mm)</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetSpacingBetweenLines(this Paragraph paragraph, double? beforeMM = null, double? afterMM = null)
        {
            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true, insertAtIdx: 0)!;

            var spacingBetweenLines = new SpacingBetweenLines();
            if (beforeMM != null) spacingBetweenLines.Before = beforeMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);
            if (afterMM != null) spacingBetweenLines.After = afterMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

            paragraphProperties.SpacingBetweenLines = spacingBetweenLines;

            return paragraph;
        }

        /// <summary>
        /// set indentation of paragraph
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="startMM">start indent (mm)</param>
        /// <param name="endMM">end indent (mm)</param>
        /// <param name="hangingMM">hanging (mm)</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetIndentation(this Paragraph paragraph,
            double? startMM = null, double? endMM = null, double? hangingMM = null)
        {
            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true, insertAtIdx: 0)!;

            var indentation = new Indentation();
            if (startMM != null) indentation.Start = startMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);
            if (endMM != null) indentation.End = endMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);
            if (hangingMM != null) indentation.Hanging = hangingMM.Value.MMToTwip().ToString(CultureInfo.InvariantCulture);

            paragraphProperties.Indentation = indentation;

            return paragraph;
        }

        /// <summary>
        /// mimic of margin using indentation(start:left, end:right) and spacing(before:top, after:bottom)
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="leftMM">left margin (mm)</param>
        /// <param name="topMM">top margin (mm)</param>
        /// <param name="rightMM">right margin (mm)</param>
        /// <param name="bottomMM">bottom margin (mm)</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetMargin(this Paragraph paragraph,
            double? leftMM = null, double? topMM = null, double? rightMM = null, double? bottomMM = null) =>
            paragraph
                .SetIndentation(startMM: leftMM, endMM: rightMM)
                .SetSpacingBetweenLines(beforeMM: topMM, afterMM: bottomMM);

        /// <summary>
        /// mimic of margin using indentation(start:left, end:right) and spacing(before:top, after:bottom)        
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="marginMM">left margin (mm)</param>        
        /// <returns>paragraph</returns>
        public static Paragraph SetUniformMargin(this Paragraph paragraph, double marginMM) =>
            paragraph.SetMargin(marginMM, marginMM, marginMM, marginMM);

        /// <summary>
        /// set paragraph justification
        /// </summary>
        /// <param name="paragraph">paragraph</param>
        /// <param name="justificationType">type of paragraph justification</param>
        /// <returns>paragraph</returns>
        public static Paragraph SetJustification(this Paragraph paragraph, JustificationValues? justificationType = null)
        {
            var paragraphProperties = paragraph.GetProperties(createIfNotExists: true, insertAtIdx: 0)!;

            if (justificationType == null)
            {
                paragraphProperties.Justification = null;
                return paragraph;
            }

            var justification = new Justification() { Val = justificationType };
            paragraphProperties.Justification = justification;

            return paragraph;
        }

    }

}