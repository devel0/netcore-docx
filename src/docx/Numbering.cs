namespace SearchAThing.DocX;

public static partial class DocXExt
{

    /// <summary>
    /// scan numbering/abstract num to retrieve max abstract num id
    /// </summary>
    internal static int GetMaxAbstractNumId(this Numbering numbering)
    {
        var q = numbering.Elements<AbstractNum>().Select(abstractNum => abstractNum.AbstractNumberId?.Value);

        if (q.Any())
        {
            return q.Max(v => v is null ? 0 : v.Value);
        }

        return 0;
    }

    public static int MaxNumberId(this WordprocessingDocument doc)
    {
        var q = doc.GetNumberingDefinitionsPart()
            .Numbering
            .Elements<NumberingInstance>()
            .Where(w => w.NumberID is not null)
            .Select(w => w.NumberID!.Value);

        if (q.Any()) return q.Max();

        return 0;
    }

    internal static AbstractNum GetAbstractNum(this WordprocessingDocument doc,
        NumberFormatValues format,
        bool structured = false,
        bool restartNumbering = false)
    {
        var numberingDefinitionsPart = doc.GetNumberingDefinitionsPart();
        var numbering = numberingDefinitionsPart.Numbering;

        var abstractNumbering = restartNumbering ? null : numbering
            .Elements<AbstractNum>()
            .Where(abstractNum =>
                abstractNum.Elements<Level>().FirstOrDefault()?.NumberingFormat?.Val?.Value == format
                &&
                abstractNum.Elements<Level>().Any(level => level.LevelText?.Val?.Value?.Contains("%1.%2") == structured)
                )
            .LastOrDefault();

        if (abstractNumbering is null)
        {
            var lib = WrapperManager.GetWrapperRef(doc);

            var newAbstractNumberingId = numbering.GetMaxAbstractNumId() + 1;

            if (format == NumberFormatValues.Bullet)
            {
                abstractNumbering = lib.GenerateAbstractNum_Bullet(newAbstractNumberingId);
            }

            else if (format == NumberFormatValues.Decimal)
            {
                if (structured)
                    abstractNumbering = lib.GenerateAbstractNum_Decimal_Structured(newAbstractNumberingId);

                else
                    abstractNumbering = lib.GenerateAbstractNum_Decimal(newAbstractNumberingId);
            }

            else if (format == NumberFormatValues.None)
            {
                abstractNumbering = lib.GenerateAbstractNum_None(newAbstractNumberingId);
            }

            else
                throw new Exception($"can't find numbering type {format}");

            int? lastAbstractNumIdx = null;
            for (int idx = 0; idx < numbering.ChildElements.Count; ++idx)
            {
                var element = numbering.ChildElements[idx];

                if (element is AbstractNum) lastAbstractNumIdx = idx;
            }

            if (lastAbstractNumIdx is not null)
                numbering.InsertAt(abstractNumbering, lastAbstractNumIdx.Value + 1);
            else
                numbering.Append(abstractNumbering);
        }

        return abstractNumbering;
    }

    internal static NumberingInstance GetNumberingInstance(this WordprocessingDocument doc,
        AbstractNum abstractNumbering)
    {
        var numberingDefinitionsPart = doc.GetNumberingDefinitionsPart();
        var numbering = numberingDefinitionsPart.Numbering;

        // var abstractNumbering = doc.GetAbstractNum(format, structured);

        var numberingInstance = numbering
            .Elements<NumberingInstance>()
            .LastOrDefault(numberingInstance =>
                numberingInstance?.AbstractNumId?.Val?.Value == abstractNumbering.AbstractNumberId?.Value);

        if (numberingInstance is null)
        {
            numberingInstance = new NumberingInstance { NumberID = doc.MaxNumberId() + 1 };
            numberingInstance.Append(new AbstractNumId { Val = abstractNumbering.AbstractNumberId });
            numbering.Append(numberingInstance);
        }

        return numberingInstance;
    }


}