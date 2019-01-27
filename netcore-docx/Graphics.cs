using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;
using SearchAThing.DocX;
using System.IO;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace SearchAThing.DocX
{

    public static class GraphicsExt
    {

        /// <summary>
        /// insert picture ( feed data )
        /// </summary>        
        public static ImagePart InsertPicture(this WordprocessingDocument doc, string pathfilename, ImagePartType type)
        {
            var mainPart = doc.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(type);
            using (var fs = new FileStream(pathfilename, FileMode.Open))
            {
                imagePart.FeedData(fs);
            }
            return imagePart;
        }

        /// <summary>
        /// get id of image part
        /// </summary>        
        public static string GetId(this ImagePart imagePart, WordprocessingDocument doc)
        {
            var mainPart = doc.MainDocumentPart;
            return mainPart.GetIdOfPart(imagePart);
        }

        /// <summary>
        /// create an inline image drawing
        /// </summary>
        public static Drawing CreateImageInlineDrawing(this ImagePart imagePart, WordprocessingDocument doc,
        double widthMM, double heightMM)
        {
            var w = (int)widthMM.MMToEMU();
            var h = (int)heightMM.MMToEMU();

            var filename = System.IO.Path.GetFileName(imagePart.Uri.OriginalString);

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = w, Cy = h },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = filename
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                          new PIC.NonVisualDrawingProperties()
                                          {
                                              Id = (UInt32Value)0U,
                                              Name = filename
                                          },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = imagePart.GetId(doc),
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = w, Cy = h }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return element;
        }

    }

}