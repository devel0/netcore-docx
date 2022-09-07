using Xunit;
using System.Linq;
using System;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace SearchAThing.DocX.Tests
{
    public partial class TestGenericTests
    {

        [Fact]
        public void TestGeneric_0001()
        {
            var outputPathfilename = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "out.docx");

            var img01pathfilename = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../examples/0001/imgs/img01.png");

            using (var doc = DocXToolkit.Create(outputPathfilename))
            {
                sampledocs.Samples.SampleDoc2(doc, img01pathfilename);
            }

        }
    }
}