using System.Linq;
using System;
using static System.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DColor = System.Drawing.Color;

using SearchAThing.DocX;
using SearchAThing;

namespace sampledocs;

public static partial class Samples
{

    /// <summary>
    /// create sample document ( used for unit test )
    /// </summary>
    public static void SampleDoc2(this WordprocessingDocument doc, string img01pathfilename)
    {
        //-----------------------------------------------------------            
        // DOC DEFAULTS, MAIN SECTION PROPERTIES, LIBRARY
        //-----------------------------------------------------------            

        doc
            .SetDocDefaults(runFontName: "Arial");

        doc
            .MainSectionProperties().SetPageSize(PaperSize.A4, PageOrientationValues.Portrait);

        var titleStyle = doc.GetPredefinedStyle(LibraryStyleEnum.Title);
        var subTitleStyle = doc.GetPredefinedStyle(LibraryStyleEnum.Subtitle);

        var heading1Style = doc.GetPredefinedStyle(LibraryStyleEnum.Heading1);
        var heading2Style = doc.GetPredefinedStyle(LibraryStyleEnum.Heading2);

        var tableStyle = doc.GetPredefinedStyle(LibraryStyleEnum.Table);
        var figureStyle = doc.GetPredefinedStyle(LibraryStyleEnum.Figure);

        //-----------------------------------------------------------            
        // TITLE, SUBTITLE
        //-----------------------------------------------------------            

        doc
            .AddParagraph("netcore-docx", titleStyle);

        doc
            .AddParagraph("Demo document", subTitleStyle);

        //-----------------------------------------------------------            
        // TOC
        //-----------------------------------------------------------            

        doc
            .AddParagraph()
            .SectionProperties()
                .SetHeader(header => header.AddParagraph("netcore-docx - Demo document"))
                .SetFooter(footer => footer
                    .AddParagraph("Page ")
                    .AddField(FieldEnum.PAGE)
                    .AddText(" of ")
                    .AddField(FieldEnum.NUMPAGES)
                    .SetJustification(JustificationValues.Center)
                );

        doc
            .AddToc("Table of Contents")
            .AddParagraph();

        doc
            .AddBreak();

        //-----------------------------------------------------------            
        // FONT FAMILY, SIZE, COLOR, PARAGRAPH PROPERTIES
        //-----------------------------------------------------------            

        doc
            .AddParagraph("Font family, size, color, paragraph properties", heading1Style)
            .EnableAutoNumbering()

            .AddParagraph("some normal paragraph")
            .AddParagraph("line with space before 5mm").SetSpacingBetweenLines(beforeMM: 5)
            .AddParagraph("line with space after 5mm").SetSpacingBetweenLines(afterMM: 5)
            .AddParagraph("bold mode").SetBold()
            .AddParagraph("italic mode").SetItalic()
            .AddParagraph("underline mode").SetUnderline()

            .AddParagraph("some color ")
                .AddRun("red", action: run => run.SetColor(DColor.Red))
                .AddRun(" green", action: run => run.SetColor(DColor.Green))
                .AddRun(" blue", action: run => run.SetColor(DColor.Blue))

            .AddParagraph("paragraph shading")
            .IncAutoNumbering()
                .AddParagraph("orange, clear").SetShading(DColor.Orange)
            .DecAutoNumbering()

            .AddParagraph("run shading")
                .IncAutoNumbering()
                .AddParagraph("clear ")
                    .AddRun("red", action: run => run.SetShading(DColor.Red, ShadingPatternValues.Clear))
                    .AddSpace().AddRun("green", action: run => run.SetShading(DColor.Green, ShadingPatternValues.Clear))
                    .AddSpace().AddRun("blue", action: run => run.SetShading(DColor.Blue, ShadingPatternValues.Clear))
                .DecAutoNumbering()

            .DisableAutoNumbering();

        doc
            .AddBreak();

        //-----------------------------------------------------------            
        // CUSTOM PARAGRAPH STYLE
        //-----------------------------------------------------------       

        doc
            .AddParagraph("Custom paragraph properties", heading1Style);

        {
            var myStyle1 = doc.AddParagraphSyle("myStyle1",
                runFontName: "Times New Roman",
                runFontColor: DColor.Blue,
                runFontSizePt: 14,
                spacingBetweenLinesOpts: new SpacingBetweenLinesOptions { AfterMM = 5 },
                indentationOpts: new IndentationOptions { },
                justification: JustificationValues.Left);

            doc.AddParagraph("paragraph with myStyle1", myStyle1);

            var myStyle2 = doc.AddParagraphSyle("myStyle2",
                runFontColor: DColor.Red,
                basedOn: myStyle1);

            doc.AddParagraph("paragraph with myStyle2 based on myStyle1", myStyle2);
        }

        //-----------------------------------------------------------            
        // FIND REPLACE
        //-----------------------------------------------------------       

        doc
            .AddParagraph("Find replace", heading1Style);

        Paragraph? pref = null;

        doc
            .AddParagraph("This ")
                .AddRun("wor", run => run.SetColor(DColor.Red))
                .AddRun("ld", run => run.SetColor(DColor.Green))
                .AddRun("s", run => run.SetColor(DColor.Blue))
                .AddRun(" text")
                .Act(p => pref = p)
            .AddParagraph($"Is composed by {pref!.GetRuns().Count()} runs")
            .AddParagraph("where the word \"worlds\" is composed by 3 runs");

        var search = doc.FindText("worlds");

        doc
            .AddParagraph(
                $"Searching the word \"worlds\" in previous text results in {search.Count} occurrences:");

        foreach (var match in search)
        {
            doc
                .AddParagraph($"paragraph: [")
                    .AddRuns(match.Paragraph.GetRuns().Select(r => (Run)r.Clone()))
                    .AddRun($"] contains {match.Runs.Count} matching runs:")
                    .SetNumbering(0, NumberFormatValues.Decimal);

            foreach (var matchingrun in match.Runs)
            {
                doc.AddParagraph()
                    .AddRun(((Run)matchingrun.Clone()).Act(run =>
                        run.SetShading(DColor.Yellow)))
                    .SetNumbering(1, NumberFormatValues.Bullet);
            }
        }

        doc
            .AddBreak();

        //-----------------------------------------------------------            
        // IMAGES
        //-----------------------------------------------------------            

        doc
            .AddParagraph("Images", heading1Style)
            .AddParagraph("Image with original size", heading2Style)
            .AddImage(img01pathfilename)
            .AddParagraph()

            .AddParagraph("Image keep aspect Width = 50mm", heading2Style)
            .AddImage(img01pathfilename, widthMM: 50)
            .AddParagraph()

            .AddParagraph("Image keep aspect Height = 50mm", heading2Style)
            .AddImage(img01pathfilename, heightMM: 50)
            .AddParagraph()

            .AddParagraph("Image unconstrained width:50mm height:50mm", heading2Style)
            .AddImage(img01pathfilename, widthMM: 50, heightMM: 50)
            .AddParagraph()

            .AddTable()
                .AddColumns(3, 1)
                .AddRow(row =>
                {
                    row.GetCell(0)
                        .SetParagraph("Left")
                        .AddImage(img01pathfilename, widthMM: 30)
                        .SetJustification(JustificationValues.Left);

                    row.GetCell(1)
                        .SetParagraph("Center")
                        .AddImage(img01pathfilename, widthMM: 30)
                        .SetJustification(JustificationValues.Center);

                    row.GetCell(2)
                        .SetParagraph("Right")
                        .AddImage(img01pathfilename, widthMM: 30)
                        .SetJustification(JustificationValues.Right);

                    var cellsCount = row.Elements<TableCell>().Count();
                    for (int i = 0; i < cellsCount; ++i) row.GetCell(i).SetPadding(2);
                })
                .SetBordersAll()
            .AddBreak();

        //-----------------------------------------------------------            
        // TABLE, FIELDS
        //-----------------------------------------------------------            

        doc
            .AddParagraph("Table with Field types", heading1Style);

        var tbl = doc
            .AddTable(tableWidthPercent: 50)
            .AddColumn(1).AddColumn(1);

        tbl.AddRow(row =>
        {
            row.GetCell(0)
                .SetShading(DColor.Navy)
                .SetParagraph("FIELD NAME", run => run.SetBold().SetColor(DColor.White))
                .SetUniformMargin(2);

            row.GetCell(1)
                .SetShading(DColor.Navy)
                .SetParagraph("SAMPLE", run => run.SetBold().SetColor(DColor.White))
                .SetUniformMargin(2);
        });

        foreach (var t in Enum.GetValues<FieldEnum>())
        {
            tbl.AddRow(row =>
            {
                row.GetCell(0).SetParagraph(t.ToString(), run => run.SetBold()).SetUniformMargin(2);
                row.GetCell(1).GetFirstChild<Paragraph>()!.AddField(t).SetUniformMargin(2);
            });
        }
        
        // short version
        //tbl.SetBordersOutside(BorderValues.Single);                

        // custom version
        tbl.SetBorders((args) =>
        {
            if (args.colIdx == 0) 
                args.leftBorder = new LeftBorder() { Val = BorderValues.Single };

            if (args.rowIdx == 0) 
                args.topBorder = new TopBorder() { Val = BorderValues.Single };

            if (args.colIdx == args.colCount - 1) 
                args.rightBorder = new RightBorder() { Val = BorderValues.Single };

            if (args.rowIdx == args.rowCount - 1) 
                args.bottomBorder = new BottomBorder() { Val = BorderValues.Single };
        });

        doc
            .AddBreak();

        //-----------------------------------------------------------            
        // NUMBERING
        //-----------------------------------------------------------            

        doc
            .AddParagraph("Numbering", heading1Style)

            .AddParagraph("Bullet list", heading2Style)
            .AddParagraph("sample1").SetNumbering(0)
            .AddParagraph("sample2").SetNumbering(1)
            .AddParagraph("sample2").SetNumbering(0)

            .AddParagraph("Decimal list", heading2Style)
            .AddParagraph("sample1").SetNumbering(0, NumberFormatValues.Decimal, restartNumbering: true)
            .AddParagraph("sample2x").SetNumbering(1, NumberFormatValues.Decimal)
            .AddParagraph("sample2").SetNumbering(0, NumberFormatValues.Decimal)

            .AddParagraph("Structured Decimal list", heading2Style)
            .AddParagraph("sample1").SetNumbering(0, NumberFormatValues.Decimal, structured: true)
            .AddParagraph("sample2").SetNumbering(1, NumberFormatValues.Decimal, structured: true)
            .AddParagraph("sample2").SetNumbering(0, NumberFormatValues.Decimal, structured: true);

        doc
            .AddBreak();

        //-----------------------------------------------------------            

        doc.Finalize();

    }

}