using System;
using System.Linq;
using static System.Math;
using System.Collections.Generic;
using System.Globalization;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DW_WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing;
using DW_PIC = DocumentFormat.OpenXml.Drawing.Pictures;

using SearchAThing;
using SearchAThing.DocX;
using static SearchAThing.DocX.Constants;
using System.IO;
using System.Security.Cryptography;

namespace SearchAThing.DocX
{

    public static partial class DocXExt
    {

        /// <summary>
        /// add image from file to the list of image parts
        /// </summary>
        /// <param name="doc">word processing document</param>
        /// <param name="imgPathfilename">image pathfilename</param>
        /// <param name="type">type of image</param>        
        /// <param name="widthMM">width (mm)</param>
        /// <param name="heightMM">height (mm)</param>        
        /// <returns>imagepart associated with given image</returns>
        public static Paragraph AddImage(this WordprocessingDocument doc,
            string imgPathfilename,
            double? widthMM = null,
            double? heightMM = null,
            ImagePartType? type = null) =>
            doc.AddParagraph("", action: run => run.AddImage(imgPathfilename, widthMM, heightMM, type));

        internal static WrapperManager.ImagePartNfo AddImagePart(this WordprocessingDocument doc,
            string imgPathfilename, ImagePartType? type = null)
        {
            var _type = ImagePartType.Bmp;
            if (type != null)
                _type = type.Value;
            else
            {
                var q = DocXToolkit.GetImagePartType(imgPathfilename);
                if (q != null) _type = q.Value;
            }

            var wrapperRef = doc.GetWrapperRef();
            var imagePartNfo = wrapperRef.TryGetImagePart(imgPathfilename);

            if (imagePartNfo is null)
            {
                var imagePart = doc.MainDocumentPart!.AddImagePart(_type);

                using (var fs = new FileStream(imgPathfilename, FileMode.Open))
                {
                    imagePart.FeedData(fs);
                }

                var nfo = wrapperRef.AddImagePartNfo(imgPathfilename, imagePart);

                return nfo;
            }
            else
                return imagePartNfo.Value;
        }

        /// <summary>
        /// add image to last run of this paragraph;
        /// if one of width, height specified other is computed maintaining aspect
        /// </summary>
        /// <param name="run">run which add image</param>
        /// <param name="imagePathfilename">pathfilename of image</param>
        /// <param name="widthMM">(optional) image width mm</param>
        /// <param name="heightMM">(optional) image height mm</param>
        /// <param name="type">(optional) image type</param>
        /// <param name="docPrId">(optional) doc property id</param>
        /// <param name="doc">(optional) if null WordprocessingDocument will retrieve from parent</param>
        /// <returns></returns>
        public static Run AddImage(this Run run, string imagePathfilename,
            double? widthMM = null, double? heightMM = null, ImagePartType? type = null, uint? docPrId = null,
            WordprocessingDocument? doc = null)
        {
            if (doc is null) doc = run.GetWordprocessingDocument();

            run.Append(doc
                .AddImagePart(imagePathfilename, type)
                .CreateImageInlineDrawing(doc, widthMM, heightMM, docPrId));

            return run;
        }

        /// <summary>
        /// create a new paragraph after this and add image;
        /// if one of width, height specified other is computed maintaining aspect
        /// </summary>
        /// <param name="paragraph">paragraph which add image</param>
        /// <param name="imagePathfilename">pathfilename of image</param>
        /// <param name="widthMM">(optional) image width mm</param>
        /// <param name="heightMM">(optional) image height mm</param>
        /// <param name="type">(optional) image type</param>
        /// <param name="docPrId">(optional) doc property id</param>
        /// <param name="doc">(optional) if null WordprocessingDocument will retrieve from parent</param>
        /// <returns>new paragraph with image inside</returns>
        public static Paragraph AddImage(this Paragraph paragraph, string imagePathfilename,
            double? widthMM = null, double? heightMM = null, ImagePartType? type = null, uint? docPrId = null,
            WordprocessingDocument? doc = null)
        {
            if (doc is null) doc = paragraph.GetWordprocessingDocument();

            return paragraph
                .AddParagraph("",
                action: run => run.AddImage(imagePathfilename, widthMM, heightMM, type, docPrId, doc));
        }

        internal static Paragraph AddInlineDrawing(this Paragraph para, Drawing drawing, Action<Run>? action = null, int? runIdx = null) =>
            para.AddRun(
                action: r => { r.AddInlineDrawing(drawing); action?.Invoke(r); },
                runIdx: runIdx);

        /// <summary>
        /// retrieve id associated with given imagepart
        /// </summary>
        /// <param name="imagePart">image part</param>
        /// <param name="doc">word processing document</param>
        public static string GetId(this ImagePart imagePart, WordprocessingDocument doc) =>
            doc.GetMainDocumentPart().GetIdOfPart(imagePart);

        /// <summary>
        /// create an inline drawing from given imagePart
        /// </summary>
        /// <param name="imagePartNfo">image part</param>
        /// <param name="doc">word processing document</param>        
        /// <param name="widthMM">width (mm)</param>
        /// <param name="heightMM">height (mm)</param>
        /// <param name="docPrId">(optional) specific doc property id</param>
        /// <returns>inline image drawing</returns>
        internal static Drawing CreateImageInlineDrawing(this WrapperManager.ImagePartNfo imagePartNfo,
            WordprocessingDocument doc,
            double? widthMM = null, double? heightMM = null,
            uint? docPrId = null)
        {
            if (docPrId is null) docPrId = doc.GetMaxDocPrId() + 1;

            var filename = System.IO.Path.GetFileName(imagePartNfo.ImagePart.Uri.OriginalString);

            var aspect = ((double)imagePartNfo.widthMM) / ((double)imagePartNfo.heightMM);

            if (widthMM is null && heightMM is not null) // w=?, h                            
                widthMM = aspect * heightMM;

            else if (widthMM is not null && heightMM is null) // w, h=?        
                heightMM = widthMM / aspect;

            else if (widthMM is null && heightMM is null) // w=?, h=?
            {
                widthMM = imagePartNfo.widthMM;
                heightMM = imagePartNfo.heightMM;
            }

            var w = widthMM!.Value.MMToEMU();
            var h = heightMM!.Value.MMToEMU();

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW_WP.Inline(
                         new DW_WP.Extent() { Cx = w, Cy = h },
                         new DW_WP.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW_WP.DocProperties()
                         {
                             Id = (uint)docPrId,
                             Name = filename
                         },
                         new DW_WP.NonVisualGraphicFrameDrawingProperties(
                             new DW.GraphicFrameLocks() { NoChangeAspect = true }),
                         new DW.Graphic(
                             new DW.GraphicData(
                                 new DW_PIC.Picture(
                                     new DW_PIC.NonVisualPictureProperties(
                                          new DW_PIC.NonVisualDrawingProperties()
                                          {
                                              Id = (UInt32Value)0U,
                                              Name = filename
                                          },
                                         new DW_PIC.NonVisualPictureDrawingProperties()),
                                     new DW_PIC.BlipFill(
                                         new DW.Blip(
                                             new DW.BlipExtensionList(
                                                 new DW.BlipExtension()
                                                 {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = imagePartNfo.ImagePart.GetId(doc),
                                             CompressionState =
                                             DW.BlipCompressionValues.Print
                                         },
                                         new DW.Stretch(
                                             new DW.FillRectangle())),
                                     new DW_PIC.ShapeProperties(
                                         new DW.Transform2D(
                                             new DW.Offset() { X = 0L, Y = 0L },
                                             new DW.Extents() { Cx = w, Cy = h }),
                                         new DW.PresetGeometry(
                                             new DW.AdjustValueList()
                                         )
                                         { Preset = DW.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U
                     });

            return element;
        }

    }

    public static partial class DocXToolkit
    {
        public static ImagePartType? GetImagePartType(string pathfilename)
        {
            var ext = Path.GetExtension(pathfilename).ToLower();
            ImagePartType? res = ext.ToLower() switch
            {
                ".bmp" => ImagePartType.Bmp,
                ".gif" => ImagePartType.Gif,
                ".png" => ImagePartType.Png,
                ".tif" or ".tiff" => ImagePartType.Tiff,
                ".ico" => ImagePartType.Icon,
                ".pcx" => ImagePartType.Pcx,
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".emf" => ImagePartType.Emf,
                ".wmf" => ImagePartType.Wmf,
                ".svg" => ImagePartType.Svg,

                _ => null
            };

            return res;
        }
    }

}