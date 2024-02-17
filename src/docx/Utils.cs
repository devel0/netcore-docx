namespace SearchAThing.DocX;

public static partial class DocXExt
{

    // COLOR

    public static string ToWPColorString(this System.Drawing.Color? color) =>
        color == null ? "auto" :
        color.Value.ToWPColorString();

    public static string ToWPColorString(this System.Drawing.Color color) =>
        color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");

    // OPENXML

    /// <summary>
    /// get idx of this element as child idx of its parent
    /// </summary>
    public static int? GetIndex(this OpenXmlElement element)
    {
        var parent = element.Parent;
        if (parent == null) return null;

        int idx = 0;
        foreach (var child in parent.ChildElements)
        {
            if (child == element) return idx;

            ++idx;
        }

        return null;
    }

    /// <summary>
    /// convert twip to mm<br/>
    /// mm = 1440/25.4 twip<br/>
    /// twip = 25.4/1440 mm
    /// </summary>
    public static double TwipToMM(this UInt32Value twip)
    {
        return ((uint)twip).TwipToMM();
    }

    /// <summary>
    /// convert 0..1 factor to fithy-thousand percent
    /// </summary>
    public static double FactorToPct(this double factor) => factor * 5000;


}

public static partial class DocXToolkit
{


    // FILE

    /// <summary>
    /// retrieve md5sum from file content
    /// </summary>
    /// <param name="pathfilename">pathfilename of file for which compute md5sum</param>
    /// <returns>file md5sum</returns>
    public static string ComputeMD5Sum(string pathfilename)
    {
        var res = "";

        using (var md5 = MD5.Create())
        {
            using (var stream = File.OpenRead(pathfilename))
            {
                var chksum = md5.ComputeHash(stream);
                res = BitConverter.ToString(chksum);
            }
        }

        return res;
    }

}