using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using DColor = System.Drawing.Color;

using SearchAThing.DocX;
using SearchAThing;

namespace examples;

class Program
{
    static void Main(string[] args)
    {
        var outputPathfilename = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "out.docx");

        var img01pathfilename = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "imgs/img01.png");

        using (var doc = DocXToolkit.Create(outputPathfilename))
        {
            sampledocs.Samples.SampleDoc2(doc, img01pathfilename);
        }

        var psi = new ProcessStartInfo(outputPathfilename);
        psi.UseShellExecute = true;
        Process.Start(psi);
    }
}
