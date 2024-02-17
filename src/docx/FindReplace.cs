namespace SearchAThing.DocX;

public static partial class DocXExt
{

    /// <summary>
    /// search for given textToSearch in all document runs ( across run supported )
    /// and replace from the first run at its matching offset ( it may not 0 )
    /// consuming run text lengths until last run that will eventually expands text length.
    /// </summary>                        
    /// <param name="doc">wordprocessing document</param>
    /// <param name="textToSearch">text to search for</param>
    /// <param name="replaceWith">text to replace on occurrences</param>
    /// <param name="splitAndCreateSingleRun">if true matching inside run will removed and first/end matching
    /// run truncated to exclude the presence of textToSearch ; a new independent run will created inside with
    /// replaceWith text all in a single run.</param>
    /// <param name="stringComparison">allow to match with case insensitive</param>
    public static List<IReadOnlyList<Run>> FindAndReplace(this WordprocessingDocument doc,
        string textToSearch, string replaceWith,
        bool splitAndCreateSingleRun = false,
        StringComparison stringComparison = StringComparison.InvariantCulture)
    {
        var res = new List<IReadOnlyList<Run>>();

        var q = doc.FindText(textToSearch, stringComparison);

        foreach (var occurrence in q)
        {
            var replace = occurrence.ReplaceText(replaceWith, splitAndCreateSingleRun);
            if (replace != null)
            {
                res.Add(replace);
            }
        }

        return res;
    }

    /// <summary>
    /// search for given textToSearch in all document runs ( across run supported );
    /// result is a list of FindTextResult each describing a paragraph runs with
    /// the first run offset from where textToSearch.Length text spans across subsequent runs.
    /// </summary>    
    public static List<FindTextResult>
    FindText(this WordprocessingDocument doc, string textToSearch,
        StringComparison stringComparison = StringComparison.InvariantCulture)
    {
        var res = new List<FindTextResult>();
        var textNfo = new TextNfo(doc);

        int off = 0;
        var text = textNfo.Text;

        while (off < text.Length)
        {
            var resRuns = new List<Run>();
            int firstRunBeginOffset = 0;

            var offFound = text.IndexOf(textToSearch, off, stringComparison);

            if (offFound == -1) break;

            var runIdx = textNfo.GetRunIdxBeginningAtOffset(offFound);

            if (runIdx != null)
            {
                var run = textNfo.Runs[runIdx.Value];
                var runBeginOffset = textNfo.RunsBeginOffset[runIdx.Value];
                var runTextLen = textNfo.RunsTextLen[runIdx.Value];

                resRuns.Add(run);
                ++runIdx;
                firstRunBeginOffset = offFound - runBeginOffset;

                var firstRunTextLength = runTextLen;

                var l = firstRunTextLength - firstRunBeginOffset;

                while (l < textToSearch.Length)
                {
                    run = textNfo.Runs[runIdx.Value];
                    l += textNfo.RunsTextLen[runIdx.Value];

                    resRuns.Add(run);
                    ++runIdx;
                }
            }

            off = offFound + textToSearch.Length;

            if (resRuns.Count > 0)
            {
                var parent = resRuns[0].Parent;

                if (resRuns.All(run => run.Parent == parent))
                    res.Add(new FindTextResult(textToSearch, resRuns, firstRunBeginOffset));
            }
        }

        return res;
    }

    /// <summary>
    /// From the search information replace searched text with new one given.
    /// </summary>        
    /// <param name="findTextResult">find text result object</param>        
    /// <param name="newText">text to replace on find text object runs</param>        
    /// <param name="splitAndCreateSingleRun">if true matching inside run will removed and first/end matching
    /// run truncated to exclude the presence of textToSearch ; a new independent run will created inside with
    /// replaceWith text all in a single run.</param>
    internal static IReadOnlyList<Run>? ReplaceText(this FindTextResult findTextResult,
        string newText,
        bool splitAndCreateSingleRun = false)
    {
        if (findTextResult.Runs.Count == 0 && newText.Length > 0) return null;

        if (splitAndCreateSingleRun)
        {
            var runs = findTextResult.Runs;

            var paragraph = findTextResult.Paragraph;

            var newRun = new Run();
            paragraph.AddText(newText, action: run => newRun = run);

            int zapped = 0;
            int required = findTextResult.TextSearch.Length;

            var insertNewRunAtIndex = findTextResult.FirstRunBeginOffset == 0 ?
                findTextResult.Runs[0].GetIndex()!.Value :
                findTextResult.Runs[0].GetIndex()!.Value + 1;

            for (int ridx = 0; ridx < findTextResult.Runs.Count; ++ridx)
            {
                var run = findTextResult.Runs[ridx];
                var runText = run.GetTextStr()!;

                if (ridx == 0)
                {
                    if (findTextResult.FirstRunBeginOffset == 0)
                    {
                        // RRRRRRRRRRRRR
                        // NNNNNNNN
                        if (findTextResult.Runs.Count == 1 && runText.Length > newText.Length)
                        {
                            run.SetText(runText.Substring(newText.Length));
                            zapped = newText.Length;
                        }

                        // RRRRRRRRRRRRR
                        // NNNNNNNNNNNNN ???
                        else
                        {
                            run.Remove();

                            zapped = runText.Length;
                        }
                    }

                    else
                    {

                        // RRRRRRRRRRRRR
                        // __NNNNNNNN
                        if (runText.Length - findTextResult.FirstRunBeginOffset > newText.Length)
                        {
                            var thisRunIdx = run.GetIndex()!;

                            Run? postRun = null;

                            // ..........RRR
                            // __NNNNNNNN
                            paragraph.AddText(
                                runText.Substring(findTextResult.FirstRunBeginOffset + newText.Length),
                                action: run => postRun = run,
                                runIdx: thisRunIdx + 1);

                            // RR
                            // __NNNNNNNN
                            run.SetText(runText.Substring(0, findTextResult.FirstRunBeginOffset));

                            run.CopyPropertiesTo(postRun!);

                            zapped = newText.Length;
                        }

                        // RRRRRRRRRRRRR
                        // __NNNNNNNNNNN ???
                        else
                        {
                            run.SetText(runText.Substring(0, findTextResult.FirstRunBeginOffset));

                            zapped = runText.Length - findTextResult.FirstRunBeginOffset;
                        }
                    }
                }

                // RRRRRRRRRRRRR RRRRRRRRRRRRR
                // ??NNNNNNNNNNN ??
                else
                {
                    if (zapped >= newText.Length) throw new InternalError($"zapped more chars than needed");

                    // RRRRRRRRRRRRR RRRRRRRRRRRRR RRRRRRRRRRRRR
                    // ??NNNNNNNNNNN NNNNNNNNNNNNN ??
                    if (ridx < findTextResult.Runs.Count - 1)
                    {
                        run.Remove();
                        zapped += runText.Length;
                    }

                    // RRRRRRRRRRRRR ... RRRRRRRRRRRRR
                    // ??NNNNNNNNNNN ... NNNNNNNNNNN??
                    else
                    {
                        var remaining = newText.Length - zapped;

                        // RRRRRRRRRRRRR ... RRRRRRRRRRRRR
                        // ??NNNNNNNNNNN ... NNNNNNNNNNNNN
                        if (runText.Length == remaining)
                        {
                            run.Remove();
                            zapped += runText.Length;
                        }

                        // RRRRRRRRRRRRR ... RRRRRRRRRRRRR
                        // ??NNNNNNNNNNN ... NNNNNNNNNNN
                        else
                        {
                            run.SetText(runText.Substring(remaining));
                            zapped += remaining;
                        }
                    }
                }
            }

            newRun.Remove();
            paragraph.InsertAt(newRun, insertNewRunAtIndex);

            return new List<Run> { newRun };
        }
        else
        {
            int placed = 0;

            for (int runIdx = 0; runIdx < findTextResult.Runs.Count; ++runIdx)
            {
                var run = findTextResult.Runs[runIdx];

                int off = 0;

                if (runIdx == 0) off = findTextResult.FirstRunBeginOffset;

                var rTxt = "";
                var runText = run.GetTextStr();

                var avail = 0;

                if (runText != null)
                {
                    avail = runText.Length - off;
                    rTxt = runText;
                }

                // place remaining on last run
                var toplace = runIdx == findTextResult.Runs.Count - 1 ?
                    (newText.Length - placed) :
                    Min(newText.Length - placed, avail);

                var runNewText =
                    rTxt.Substring(0, off) +
                    newText.Substring(placed, toplace) +
                    rTxt.Substring(off + toplace);

                run.SetText(runNewText);

                placed += toplace;
            }

            return findTextResult.Runs;
        }
    }

}