using System.Linq;
using System;
using static System.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using SearchAThing.DocX;

namespace sampledocs;
public static partial class Samples
{

    /// <summary>
    /// create sample document ( used for unit test )
    /// </summary>
    public static void SampleDoc1(this WordprocessingDocument doc)
    {
        var style_My1 = doc.AddParagraphSyle("My1",
            runFontName: "Arial",
            justification: JustificationValues.Center);

        var style_My2 = doc.AddParagraphSyle("My2",
            runFontName: "Arial",
            runFontColor: System.Drawing.Color.Green,
            justification: JustificationValues.Center);

        var p = doc.AddParagraph(style: style_My1)
            .AddText("xxx", action: run => run.SetColor(System.Drawing.Color.Green))
            .AddText("Sample", action: run => run
                .SetUnderline(UnderlineValues.Single)
                .SetColor(System.Drawing.Color.Blue))
            .AddText("1", action: run => run.SetBold(true).SetItalic(true));

        doc.AddParagraph()
            .SetText("Sample2", action: run => run
                .SetFontName("MonoSpace"))
            .SetParagraphStyle(style_My2)
            .SetSpacingBetweenLines(beforeMM: 10, afterMM: 20)

            .AddText("BeforeSampleAfter", action: run => run
                .SetFontName("Lato")
                .SetUnderline()
                .SetColor(System.Drawing.Color.Red))
                .SetJustification(JustificationValues.Left);

        doc.AddParagraph()
            .AddText("aSam", action: run => run.SetColor(System.Drawing.Color.Red))
            .AddText("pl", action: run => run.SetColor(System.Drawing.Color.Green))
            .AddText("er", action: run => run.SetColor(System.Drawing.Color.Blue))
            .AddText("3");

        doc.AddParagraph()
            .AddText("aSam", action: run => run.SetColor(System.Drawing.Color.Red))
            .AddText("pl", action: run => run.SetColor(System.Drawing.Color.Green))
            .AddText("e", action: run => run.SetColor(System.Drawing.Color.Blue));
    }

}
