using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;
using SearchAThing.DocX;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace SearchAThing.DocX
{

    public static class UtilExt
    {

        /// <summary>
        /// convert mm to twip
        /// twip = 1/20 pp = 1/20 * 1/72 in = 1/1440 in = 1/1440 * 25.4 mm
        /// mm = 1440/25.4 twip
        /// </summary>        
        public static int MMToTwip(this double mm)
        {
            return (int)Round(1440d / 25.4 * mm);
        }

        /// <summary>
        /// convert mm to twip
        /// twip = 1/20 pp = 1/20 * 1/72 in = 1/1440 in = 1/1440 * 25.4 mm
        /// mm = 1440/25.4 twip
        /// </summary>        
        public static int MMToTwip(this int mm)
        {
            return ((double)mm).MMToTwip();
        }

        /// <summary>        
        /// 10mm = 360000 EMU
        /// mm = 36000 EMU
        /// </summary>
        public static int MMToEMU(this int mm)
        {
            return mm * 36000;
        }

        /// <summary>        
        /// 10mm = 360000 EMU
        /// mm = 36000 EMU
        /// </summary>
        public static int MMToEMU(this double mm)
        {
            return (int)Round(mm * 36000, 0);
        }

        /// <summary>
        /// convert twip to mm
        /// mm = 1440/25.4 twip
        /// twip = 25.4/1440 mm
        /// </summary>
        public static double TwipToMM(this uint twip)
        {
            return 25.4 / 1440 * twip;
        }

        /// <summary>
        /// convert twip to mm
        /// mm = 1440/25.4 twip
        /// twip = 25.4/1440 mm
        /// </summary>
        public static double TwipToMM(this int twip)
        {
            return ((uint)twip).TwipToMM();
        }

        /// <summary>
        /// convert given percent 0..100 to fiftieths of a Percent
        /// </summary>
        public static int Pct(this double percent)
        {
            return (int)(percent * 50);
        }

        /// <summary>
        /// walk through descendants of given type to retrieve idx-th element
        /// </summary>
        public static T DescendantAt<T>(this OpenXmlElement el, int idx) where T : OpenXmlElement
        {
            var x = el.Descendants<T>();
            int i = 0;
            foreach (var y in x)
            {
                if (i++ == idx) return y;
            }
            return null;
        }

        /// <summary>
        /// body part of MainDocumentPart.Document
        /// </summary>
        public static Body Body(this WordprocessingDocument doc)
        {
            return doc.MainDocumentPart.Document.Descendants<Body>().FirstOrDefault();
        }

        /// <summary>
        /// retrieve max id of DocProperties
        /// </summary>
        public static uint MaxDocPrId(this WordprocessingDocument doc)
        {
            return doc
               .MainDocumentPart
               .RootElement
               .Descendants<DocProperties>()
               .Max(x => (uint?)x.Id) ?? 0;
        }

    }


}
