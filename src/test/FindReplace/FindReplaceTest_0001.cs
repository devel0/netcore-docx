using Xunit;
using System.Linq;
using System;
using System.Diagnostics;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace SearchAThing.DocX.Tests
{
    public partial class FindReplaceTests
    {

        [Fact]
        public void FindReplaceTest_0001()
        {
            var outputFilename = "out.docx";
            var openWhenFinished = false;            
            
            using (var doc = DocXToolkit.Create(outputFilename))
            {      
                sampledocs.Samples.SampleDoc1(doc);

                var xmlBefore = doc.DocumentOuterXML();

                doc.FindAndReplace("Sample", "SOMPLE", splitAndCreateSingleRun: true);

                var xmlAfter = doc.DocumentOuterXML();

                Assert.Equal(File.ReadAllText("FindReplace/FindReplaceTest_0001-before.xml"), xmlBefore);
                Assert.Equal(File.ReadAllText("FindReplace/FindReplaceTest_0001-after.xml"), xmlAfter);

            }

            if (openWhenFinished)
            {
                var psi = new ProcessStartInfo(outputFilename);
                psi.UseShellExecute = true;
                Process.Start(psi);
            }
        }
    }
}