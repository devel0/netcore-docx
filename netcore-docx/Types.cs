using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using static System.Math;
using SearchAThing.DocX;
using System.Collections.Generic;
using System.Text;

namespace SearchAThing.DocX
{

    public class WPDocumentCreateParams
    {
        public PaperSize paperSize { get; set; } = PaperSize.A4;
        public PageOrientationValues orientation { get; set; } = PageOrientationValues.Portrait;
        public double marginLeftMM { get; set; } = 20;
        public double marginTopMM { get; set; } = 20;
        public double marginRightMM { get; set; } = 20;
        public double marginBottomMM { get; set; } = 20;
    };

    public class SpacingBetweenLinesOptions
    {
        public double? BeforeMM { get; set; } = null;
        public double? AfterMM { get; set; } = null;
        public double? LineHeightMM { get; set; } = null;
    };

    public class IndentationOptions
    {
        public double? StartMM { get; set; } = null;
        public double? EndMM { get; set; } = null;
        public double? HangingMM { get; set; } = null;
    };

    /// <summary>
    /// ECMA-376-1:2016 - 17.16.1 Syntax
    /// </summary>
    public enum FieldEnum
    {
        // date and time
        CREATEDATE,
        DATE,
        EDITTIME,
        PRINTDATE,
        SAVEDATE,
        TIME,

        // document information
        FILENAME,
        FILESIZE,
        LASTSAVEDBY,
        NUMCHARS,
        NUMPAGES,
        NUMWORDS,
        TEMPLATE,        

        // numbering        
        PAGE,
        REVNUM,

    };

}
