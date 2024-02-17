namespace SearchAThing.DocX;

/// <summary>
/// Spacing between lines optional arguments
/// </summary>
public class SpacingBetweenLinesOptions
{
    public double? BeforeMM { get; set; } = null;
    public double? AfterMM { get; set; } = null;
    public double? LineHeightMM { get; set; } = null;
};

/// <summary>
/// Indentation optional arguments
/// </summary>
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