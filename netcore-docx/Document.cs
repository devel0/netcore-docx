using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;
using SearchAThing.DocX;

namespace SearchAThing.DocX
{

    public static class DocExt
    {

        /// <summary>
        /// create new empty doc
        /// </summary>
        public static void Create(string pathfilename,
            PaperType paperType = PaperType.A4,
            PageOrientationValues orientation = PageOrientationValues.Portrait,
            double marginLeftMM = 0,
            double marginTopMM = 0,
            double marginRightMM = 0,
            double marginBottomMM = 0)
        {
            using (var doc = WordprocessingDocument.Create(pathfilename, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document();

                var body = main.Document.AppendChild(new Body());

                /*var p = body.AppendChild(new Paragraph());
                var r = p.AppendChild(new Run());
                r.AppendChild(new Text(DateTime.Now.ToString()));*/

                var sect = body.AppendChild(new SectionProperties());
                sect.SetPaper(paperType);
                sect.SetOrientation(orientation);
                sect.SetMargin(marginLeftMM, marginTopMM, marginRightMM, marginBottomMM);

                doc.Save();
            }
        }       

    }


}
