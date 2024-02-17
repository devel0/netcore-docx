namespace SearchAThing.DocX;

/// <summary>
/// helper class to map run offsets
/// </summary>
public class TextNfo
{
    List<Run> runs = new List<Run>();
    public IReadOnlyList<Run> Runs => runs;

    List<int> runsTextLen = new List<int>();
    public IReadOnlyList<int> RunsTextLen => runsTextLen;

    List<int> runBeginOffset = new List<int>();
    public IReadOnlyList<int> RunsBeginOffset => runBeginOffset;

    public string Text { get; private set; }

    public TextNfo(WordprocessingDocument doc)
    {
        var sb = new StringBuilder();

        int offset = 0;

        foreach (var paragraph in doc.GetParagraphs())
        {

            foreach (var run in paragraph.GetRuns())
            {
                var txt = run.GetTextStr();

                if (txt != null)
                {
                    runsTextLen.Add(txt.Length);
                    runBeginOffset.Add(offset);
                    sb.Append(txt);
                    offset += txt.Length;
                }
                else
                {
                    runsTextLen.Add(0);
                    runBeginOffset.Add(offset);
                }

                runs.Add(run);

            }

            sb.AppendLine();

            offset += Environment.NewLine.Length;
        }

        Text = sb.ToString();
    }

    public int? GetRunIdxBeginningAtOffset(int offset)
    {
        int runIdx = 0;

        foreach (var run in runs)
        {
            var _runBeginOffset = runBeginOffset[runIdx];
            var _runTextLen = runsTextLen[runIdx];

            if (_runBeginOffset + _runTextLen > offset) return runIdx;

            ++runIdx;
        }

        return null;
    }


};

/// <summary>
/// contains results of FindText information about which runs contains TextSearch
/// with information about the offset for the first run in the FirstRunBeginOffset
/// </summary>
public class FindTextResult
{

    /// <summary>
    /// runs (with same parent paragraph);
    /// TextSearch occurrence start at first run with FirstRunBeginOffset
    /// and continue to following Runs for remaining length of the TextSearch
    /// </summary>        
    public IReadOnlyList<Run> Runs { get; private set; }
    public int FirstRunBeginOffset { get; private set; }
    public string TextSearch { get; private set; }
    public Paragraph Paragraph { get; private set; }

    public FindTextResult(string textSearch, IReadOnlyList<Run> runs, int firstRunBeginOffset)
    {
        TextSearch = textSearch;
        Runs = runs;
        FirstRunBeginOffset = firstRunBeginOffset;
        Paragraph = (Paragraph)(Runs[0].Parent!);
    }

};