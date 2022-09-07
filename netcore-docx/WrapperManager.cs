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


    }

}