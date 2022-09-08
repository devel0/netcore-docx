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
using System.Runtime.CompilerServices;
using System.IO;

namespace SearchAThing.DocX
{

    public partial class WrapperManager
    {


        internal WordprocessingDocument doc;

        public WrapperManager(WordprocessingDocument doc)
        {
            this.doc = doc;
        }

        static Dictionary<WordprocessingDocument, WrapperManager> docToWrapperInstances =
            new Dictionary<WordprocessingDocument, WrapperManager>();

        internal static WrapperManager GetWrapperRef(WordprocessingDocument doc)
        {
            if (!docToWrapperInstances.TryGetValue(doc, out var lib))
            {
                lib = new WrapperManager(doc);
                docToWrapperInstances.Add(doc, lib);
            }
            return lib;
        }

        public static void Release(WordprocessingDocument doc)
        {
            if (docToWrapperInstances.TryGetValue(doc, out var lib))
            {
                docToWrapperInstances.Remove(doc);
            }
        }

        internal NumberFormatValues? numberingEnabled = null;

        internal int numberingLevel = 0;

        internal bool numberingStructured = false;

        #region inserted image cache

        public record struct ImagePartNfo
        {
            public ImagePart ImagePart;
            public string KeyNfo;
            public string Md5Sum;

            public double widthMM;
            public double heightMM;
        }

        internal Dictionary<string, List<ImagePartNfo>> inserted_images = new Dictionary<string, List<ImagePartNfo>>();

        string ImagePartKeyNfo(string imagePathfilename)
        {
            var filenfo = new FileInfo(imagePathfilename);
            return $"{imagePathfilename}_{filenfo.Length}";
        }

        internal ImagePartNfo? TryGetImagePart(string imagePathfilename)
        {
            var keyNfo = ImagePartKeyNfo(imagePathfilename);

            if (inserted_images.TryGetValue(keyNfo, out var imagePartNfoList))
            {
                if (imagePartNfoList.Count == 1)
                    return imagePartNfoList.First();

                var this_md5sum = DocXToolkit.ComputeMD5Sum(imagePathfilename);

                var imgNfo = UtilToolkit.GetImageNfo(imagePathfilename);

                var imgSize = imgNfo.ImageSizeMM();

                var q = imagePartNfoList.First(w => w.Md5Sum == this_md5sum);

                return q;
            }

            return null;
        }

        /// <summary>
        /// call this is ensured trygetimagepart didn't find image
        /// </summary>        
        internal ImagePartNfo AddImagePartNfo(string pathfilename, ImagePart imagePart)
        {
            ImagePartNfo res;

            var keyNfo = ImagePartKeyNfo(pathfilename);

            if (!inserted_images.TryGetValue(keyNfo, out var lst))
            {
                lst = new List<ImagePartNfo>();
                inserted_images.Add(keyNfo, lst);
            }

            var imgNfo = UtilToolkit.GetImageNfo(pathfilename);

            var imgSize = imgNfo.ImageSizeMM();

            var md5sum = DocXToolkit.ComputeMD5Sum(pathfilename);

            res = new ImagePartNfo
            {
                ImagePart = imagePart,
                Md5Sum = md5sum,
                widthMM = imgSize.widthMM,
                heightMM = imgSize.heightMM
            };

            lst.Add(res);

            return res;
        }

        #endregion


    }

}