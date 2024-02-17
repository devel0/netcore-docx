
using DocumentFormat.OpenXml;

using SearchAThing;
using SearchAThing.DocX;
using static SearchAThing.DocX.Constants;
using System.Runtime.CompilerServices;

namespace SearchAThing.DocX;

public partial class WrapperManager
{

    public AbstractNum GenerateAbstractNum_Bullet(int abstractNumId)
    {
        AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = abstractNumId };

        Level level1 = new Level() { LevelIndex = 0 };
        StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText1 = new LevelText() { Val = "·" };
        LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

        Tabs tabs1 = new Tabs();
        TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

        tabs1.Append(tabStop1);
        Indentation indentation1 = new Indentation() { Left = "720", Hanging = "360" };

        previousParagraphProperties1.Append(tabs1);
        previousParagraphProperties1.Append(indentation1);

        NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
        RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol", ComplexScript = "Symbol" };

        numberingSymbolRunProperties1.Append(runFonts1);

        level1.Append(startNumberingValue1);
        level1.Append(numberingFormat1);
        level1.Append(levelText1);
        level1.Append(levelJustification1);
        level1.Append(previousParagraphProperties1);
        level1.Append(numberingSymbolRunProperties1);

        Level level2 = new Level() { LevelIndex = 1 };
        StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText2 = new LevelText() { Val = "◦" };
        LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

        Tabs tabs2 = new Tabs();
        TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 1080 };

        tabs2.Append(tabStop2);
        Indentation indentation2 = new Indentation() { Left = "1080", Hanging = "360" };

        previousParagraphProperties2.Append(tabs2);
        previousParagraphProperties2.Append(indentation2);

        NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
        RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "OpenSymbol", HighAnsi = "OpenSymbol", ComplexScript = "OpenSymbol" };

        numberingSymbolRunProperties2.Append(runFonts2);

        level2.Append(startNumberingValue2);
        level2.Append(numberingFormat2);
        level2.Append(levelText2);
        level2.Append(levelJustification2);
        level2.Append(previousParagraphProperties2);
        level2.Append(numberingSymbolRunProperties2);

        Level level3 = new Level() { LevelIndex = 2 };
        StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText3 = new LevelText() { Val = "▪" };
        LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

        Tabs tabs3 = new Tabs();
        TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

        tabs3.Append(tabStop3);
        Indentation indentation3 = new Indentation() { Left = "1440", Hanging = "360" };

        previousParagraphProperties3.Append(tabs3);
        previousParagraphProperties3.Append(indentation3);

        NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
        RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "OpenSymbol", HighAnsi = "OpenSymbol", ComplexScript = "OpenSymbol" };

        numberingSymbolRunProperties3.Append(runFonts3);

        level3.Append(startNumberingValue3);
        level3.Append(numberingFormat3);
        level3.Append(levelText3);
        level3.Append(levelJustification3);
        level3.Append(previousParagraphProperties3);
        level3.Append(numberingSymbolRunProperties3);

        Level level4 = new Level() { LevelIndex = 3 };
        StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText4 = new LevelText() { Val = "·" };
        LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

        Tabs tabs4 = new Tabs();
        TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 1800 };

        tabs4.Append(tabStop4);
        Indentation indentation4 = new Indentation() { Left = "1800", Hanging = "360" };

        previousParagraphProperties4.Append(tabs4);
        previousParagraphProperties4.Append(indentation4);

        NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
        RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol", ComplexScript = "Symbol" };

        numberingSymbolRunProperties4.Append(runFonts4);

        level4.Append(startNumberingValue4);
        level4.Append(numberingFormat4);
        level4.Append(levelText4);
        level4.Append(levelJustification4);
        level4.Append(previousParagraphProperties4);
        level4.Append(numberingSymbolRunProperties4);

        Level level5 = new Level() { LevelIndex = 4 };
        StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText5 = new LevelText() { Val = "◦" };
        LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

        Tabs tabs5 = new Tabs();
        TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

        tabs5.Append(tabStop5);
        Indentation indentation5 = new Indentation() { Left = "2160", Hanging = "360" };

        previousParagraphProperties5.Append(tabs5);
        previousParagraphProperties5.Append(indentation5);

        NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
        RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "OpenSymbol", HighAnsi = "OpenSymbol", ComplexScript = "OpenSymbol" };

        numberingSymbolRunProperties5.Append(runFonts5);

        level5.Append(startNumberingValue5);
        level5.Append(numberingFormat5);
        level5.Append(levelText5);
        level5.Append(levelJustification5);
        level5.Append(previousParagraphProperties5);
        level5.Append(numberingSymbolRunProperties5);

        Level level6 = new Level() { LevelIndex = 5 };
        StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText6 = new LevelText() { Val = "▪" };
        LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

        Tabs tabs6 = new Tabs();
        TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 2520 };

        tabs6.Append(tabStop6);
        Indentation indentation6 = new Indentation() { Left = "2520", Hanging = "360" };

        previousParagraphProperties6.Append(tabs6);
        previousParagraphProperties6.Append(indentation6);

        NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
        RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "OpenSymbol", HighAnsi = "OpenSymbol", ComplexScript = "OpenSymbol" };

        numberingSymbolRunProperties6.Append(runFonts6);

        level6.Append(startNumberingValue6);
        level6.Append(numberingFormat6);
        level6.Append(levelText6);
        level6.Append(levelJustification6);
        level6.Append(previousParagraphProperties6);
        level6.Append(numberingSymbolRunProperties6);

        Level level7 = new Level() { LevelIndex = 6 };
        StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText7 = new LevelText() { Val = "·" };
        LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

        Tabs tabs7 = new Tabs();
        TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

        tabs7.Append(tabStop7);
        Indentation indentation7 = new Indentation() { Left = "2880", Hanging = "360" };

        previousParagraphProperties7.Append(tabs7);
        previousParagraphProperties7.Append(indentation7);

        NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
        RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Symbol", HighAnsi = "Symbol", ComplexScript = "Symbol" };

        numberingSymbolRunProperties7.Append(runFonts7);

        level7.Append(startNumberingValue7);
        level7.Append(numberingFormat7);
        level7.Append(levelText7);
        level7.Append(levelJustification7);
        level7.Append(previousParagraphProperties7);
        level7.Append(numberingSymbolRunProperties7);

        Level level8 = new Level() { LevelIndex = 7 };
        StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText8 = new LevelText() { Val = "◦" };
        LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

        Tabs tabs8 = new Tabs();
        TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 3240 };

        tabs8.Append(tabStop8);
        Indentation indentation8 = new Indentation() { Left = "3240", Hanging = "360" };

        previousParagraphProperties8.Append(tabs8);
        previousParagraphProperties8.Append(indentation8);

        NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
        RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "OpenSymbol", HighAnsi = "OpenSymbol", ComplexScript = "OpenSymbol" };

        numberingSymbolRunProperties8.Append(runFonts8);

        level8.Append(startNumberingValue8);
        level8.Append(numberingFormat8);
        level8.Append(levelText8);
        level8.Append(levelJustification8);
        level8.Append(previousParagraphProperties8);
        level8.Append(numberingSymbolRunProperties8);

        Level level9 = new Level() { LevelIndex = 8 };
        StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
        LevelText levelText9 = new LevelText() { Val = "▪" };
        LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

        Tabs tabs9 = new Tabs();
        TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

        tabs9.Append(tabStop9);
        Indentation indentation9 = new Indentation() { Left = "3600", Hanging = "360" };

        previousParagraphProperties9.Append(tabs9);
        previousParagraphProperties9.Append(indentation9);

        NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
        RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "OpenSymbol", HighAnsi = "OpenSymbol", ComplexScript = "OpenSymbol" };

        numberingSymbolRunProperties9.Append(runFonts9);

        level9.Append(startNumberingValue9);
        level9.Append(numberingFormat9);
        level9.Append(levelText9);
        level9.Append(levelJustification9);
        level9.Append(previousParagraphProperties9);
        level9.Append(numberingSymbolRunProperties9);

        abstractNum1.Append(level1);
        abstractNum1.Append(level2);
        abstractNum1.Append(level3);
        abstractNum1.Append(level4);
        abstractNum1.Append(level5);
        abstractNum1.Append(level6);
        abstractNum1.Append(level7);
        abstractNum1.Append(level8);
        abstractNum1.Append(level9);
        return abstractNum1;
    }

    public AbstractNum GenerateAbstractNum_Decimal(int abstractNumId)
    {
        AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = abstractNumId };

        Level level1 = new Level() { LevelIndex = 0 };
        StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText1 = new LevelText() { Val = "%1." };
        LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

        Tabs tabs1 = new Tabs();
        TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

        tabs1.Append(tabStop1);
        Indentation indentation1 = new Indentation() { Left = "720", Hanging = "360" };

        previousParagraphProperties1.Append(tabs1);
        previousParagraphProperties1.Append(indentation1);
        NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();

        level1.Append(startNumberingValue1);
        level1.Append(numberingFormat1);
        level1.Append(levelText1);
        level1.Append(levelJustification1);
        level1.Append(previousParagraphProperties1);
        level1.Append(numberingSymbolRunProperties1);

        Level level2 = new Level() { LevelIndex = 1 };
        StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText2 = new LevelText() { Val = "%2." };
        LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

        Tabs tabs2 = new Tabs();
        TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 1080 };

        tabs2.Append(tabStop2);
        Indentation indentation2 = new Indentation() { Left = "1080", Hanging = "360" };

        previousParagraphProperties2.Append(tabs2);
        previousParagraphProperties2.Append(indentation2);
        NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();

        level2.Append(startNumberingValue2);
        level2.Append(numberingFormat2);
        level2.Append(levelText2);
        level2.Append(levelJustification2);
        level2.Append(previousParagraphProperties2);
        level2.Append(numberingSymbolRunProperties2);

        Level level3 = new Level() { LevelIndex = 2 };
        StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText3 = new LevelText() { Val = "%3." };
        LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

        Tabs tabs3 = new Tabs();
        TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

        tabs3.Append(tabStop3);
        Indentation indentation3 = new Indentation() { Left = "1440", Hanging = "360" };

        previousParagraphProperties3.Append(tabs3);
        previousParagraphProperties3.Append(indentation3);
        NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();

        level3.Append(startNumberingValue3);
        level3.Append(numberingFormat3);
        level3.Append(levelText3);
        level3.Append(levelJustification3);
        level3.Append(previousParagraphProperties3);
        level3.Append(numberingSymbolRunProperties3);

        Level level4 = new Level() { LevelIndex = 3 };
        StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText4 = new LevelText() { Val = "%4." };
        LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

        Tabs tabs4 = new Tabs();
        TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 1800 };

        tabs4.Append(tabStop4);
        Indentation indentation4 = new Indentation() { Left = "1800", Hanging = "360" };

        previousParagraphProperties4.Append(tabs4);
        previousParagraphProperties4.Append(indentation4);
        NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();

        level4.Append(startNumberingValue4);
        level4.Append(numberingFormat4);
        level4.Append(levelText4);
        level4.Append(levelJustification4);
        level4.Append(previousParagraphProperties4);
        level4.Append(numberingSymbolRunProperties4);

        Level level5 = new Level() { LevelIndex = 4 };
        StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText5 = new LevelText() { Val = "%5." };
        LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

        Tabs tabs5 = new Tabs();
        TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

        tabs5.Append(tabStop5);
        Indentation indentation5 = new Indentation() { Left = "2160", Hanging = "360" };

        previousParagraphProperties5.Append(tabs5);
        previousParagraphProperties5.Append(indentation5);
        NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();

        level5.Append(startNumberingValue5);
        level5.Append(numberingFormat5);
        level5.Append(levelText5);
        level5.Append(levelJustification5);
        level5.Append(previousParagraphProperties5);
        level5.Append(numberingSymbolRunProperties5);

        Level level6 = new Level() { LevelIndex = 5 };
        StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText6 = new LevelText() { Val = "%6." };
        LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

        Tabs tabs6 = new Tabs();
        TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 2520 };

        tabs6.Append(tabStop6);
        Indentation indentation6 = new Indentation() { Left = "2520", Hanging = "360" };

        previousParagraphProperties6.Append(tabs6);
        previousParagraphProperties6.Append(indentation6);
        NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();

        level6.Append(startNumberingValue6);
        level6.Append(numberingFormat6);
        level6.Append(levelText6);
        level6.Append(levelJustification6);
        level6.Append(previousParagraphProperties6);
        level6.Append(numberingSymbolRunProperties6);

        Level level7 = new Level() { LevelIndex = 6 };
        StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText7 = new LevelText() { Val = "%7." };
        LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

        Tabs tabs7 = new Tabs();
        TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

        tabs7.Append(tabStop7);
        Indentation indentation7 = new Indentation() { Left = "2880", Hanging = "360" };

        previousParagraphProperties7.Append(tabs7);
        previousParagraphProperties7.Append(indentation7);
        NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();

        level7.Append(startNumberingValue7);
        level7.Append(numberingFormat7);
        level7.Append(levelText7);
        level7.Append(levelJustification7);
        level7.Append(previousParagraphProperties7);
        level7.Append(numberingSymbolRunProperties7);

        Level level8 = new Level() { LevelIndex = 7 };
        StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText8 = new LevelText() { Val = "%8." };
        LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

        Tabs tabs8 = new Tabs();
        TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 3240 };

        tabs8.Append(tabStop8);
        Indentation indentation8 = new Indentation() { Left = "3240", Hanging = "360" };

        previousParagraphProperties8.Append(tabs8);
        previousParagraphProperties8.Append(indentation8);
        NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();

        level8.Append(startNumberingValue8);
        level8.Append(numberingFormat8);
        level8.Append(levelText8);
        level8.Append(levelJustification8);
        level8.Append(previousParagraphProperties8);
        level8.Append(numberingSymbolRunProperties8);

        Level level9 = new Level() { LevelIndex = 8 };
        StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText9 = new LevelText() { Val = "%9." };
        LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

        Tabs tabs9 = new Tabs();
        TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

        tabs9.Append(tabStop9);
        Indentation indentation9 = new Indentation() { Left = "3600", Hanging = "360" };

        previousParagraphProperties9.Append(tabs9);
        previousParagraphProperties9.Append(indentation9);
        NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();

        level9.Append(startNumberingValue9);
        level9.Append(numberingFormat9);
        level9.Append(levelText9);
        level9.Append(levelJustification9);
        level9.Append(previousParagraphProperties9);
        level9.Append(numberingSymbolRunProperties9);

        abstractNum1.Append(level1);
        abstractNum1.Append(level2);
        abstractNum1.Append(level3);
        abstractNum1.Append(level4);
        abstractNum1.Append(level5);
        abstractNum1.Append(level6);
        abstractNum1.Append(level7);
        abstractNum1.Append(level8);
        abstractNum1.Append(level9);
        return abstractNum1;
    }

    public AbstractNum GenerateAbstractNum_None(int abstractNumId)
    {
        AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = abstractNumId };

        Level level1 = new Level() { LevelIndex = 0 };
        StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix1 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText1 = new LevelText() { Val = "" };
        LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

        Tabs tabs1 = new Tabs();
        TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs1.Append(tabStop1);
        Indentation indentation1 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties1.Append(tabs1);
        previousParagraphProperties1.Append(indentation1);

        level1.Append(startNumberingValue1);
        level1.Append(numberingFormat1);
        level1.Append(levelSuffix1);
        level1.Append(levelText1);
        level1.Append(levelJustification1);
        level1.Append(previousParagraphProperties1);

        Level level2 = new Level() { LevelIndex = 1 };
        StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix2 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText2 = new LevelText() { Val = "" };
        LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

        Tabs tabs2 = new Tabs();
        TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs2.Append(tabStop2);
        Indentation indentation2 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties2.Append(tabs2);
        previousParagraphProperties2.Append(indentation2);

        level2.Append(startNumberingValue2);
        level2.Append(numberingFormat2);
        level2.Append(levelSuffix2);
        level2.Append(levelText2);
        level2.Append(levelJustification2);
        level2.Append(previousParagraphProperties2);

        Level level3 = new Level() { LevelIndex = 2 };
        StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix3 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText3 = new LevelText() { Val = "" };
        LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

        Tabs tabs3 = new Tabs();
        TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs3.Append(tabStop3);
        Indentation indentation3 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties3.Append(tabs3);
        previousParagraphProperties3.Append(indentation3);

        level3.Append(startNumberingValue3);
        level3.Append(numberingFormat3);
        level3.Append(levelSuffix3);
        level3.Append(levelText3);
        level3.Append(levelJustification3);
        level3.Append(previousParagraphProperties3);

        Level level4 = new Level() { LevelIndex = 3 };
        StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix4 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText4 = new LevelText() { Val = "" };
        LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

        Tabs tabs4 = new Tabs();
        TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs4.Append(tabStop4);
        Indentation indentation4 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties4.Append(tabs4);
        previousParagraphProperties4.Append(indentation4);

        level4.Append(startNumberingValue4);
        level4.Append(numberingFormat4);
        level4.Append(levelSuffix4);
        level4.Append(levelText4);
        level4.Append(levelJustification4);
        level4.Append(previousParagraphProperties4);

        Level level5 = new Level() { LevelIndex = 4 };
        StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix5 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText5 = new LevelText() { Val = "" };
        LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

        Tabs tabs5 = new Tabs();
        TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs5.Append(tabStop5);
        Indentation indentation5 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties5.Append(tabs5);
        previousParagraphProperties5.Append(indentation5);

        level5.Append(startNumberingValue5);
        level5.Append(numberingFormat5);
        level5.Append(levelSuffix5);
        level5.Append(levelText5);
        level5.Append(levelJustification5);
        level5.Append(previousParagraphProperties5);

        Level level6 = new Level() { LevelIndex = 5 };
        StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix6 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText6 = new LevelText() { Val = "" };
        LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

        Tabs tabs6 = new Tabs();
        TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs6.Append(tabStop6);
        Indentation indentation6 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties6.Append(tabs6);
        previousParagraphProperties6.Append(indentation6);

        level6.Append(startNumberingValue6);
        level6.Append(numberingFormat6);
        level6.Append(levelSuffix6);
        level6.Append(levelText6);
        level6.Append(levelJustification6);
        level6.Append(previousParagraphProperties6);

        Level level7 = new Level() { LevelIndex = 6 };
        StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix7 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText7 = new LevelText() { Val = "" };
        LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

        Tabs tabs7 = new Tabs();
        TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs7.Append(tabStop7);
        Indentation indentation7 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties7.Append(tabs7);
        previousParagraphProperties7.Append(indentation7);

        level7.Append(startNumberingValue7);
        level7.Append(numberingFormat7);
        level7.Append(levelSuffix7);
        level7.Append(levelText7);
        level7.Append(levelJustification7);
        level7.Append(previousParagraphProperties7);

        Level level8 = new Level() { LevelIndex = 7 };
        StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix8 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText8 = new LevelText() { Val = "" };
        LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

        Tabs tabs8 = new Tabs();
        TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs8.Append(tabStop8);
        Indentation indentation8 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties8.Append(tabs8);
        previousParagraphProperties8.Append(indentation8);

        level8.Append(startNumberingValue8);
        level8.Append(numberingFormat8);
        level8.Append(levelSuffix8);
        level8.Append(levelText8);
        level8.Append(levelJustification8);
        level8.Append(previousParagraphProperties8);

        Level level9 = new Level() { LevelIndex = 8 };
        StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.None };
        LevelSuffix levelSuffix9 = new LevelSuffix() { Val = LevelSuffixValues.Nothing };
        LevelText levelText9 = new LevelText() { Val = "" };
        LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

        Tabs tabs9 = new Tabs();
        TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 0 };

        tabs9.Append(tabStop9);
        Indentation indentation9 = new Indentation() { Left = "0", Hanging = "0" };

        previousParagraphProperties9.Append(tabs9);
        previousParagraphProperties9.Append(indentation9);

        level9.Append(startNumberingValue9);
        level9.Append(numberingFormat9);
        level9.Append(levelSuffix9);
        level9.Append(levelText9);
        level9.Append(levelJustification9);
        level9.Append(previousParagraphProperties9);

        abstractNum1.Append(level1);
        abstractNum1.Append(level2);
        abstractNum1.Append(level3);
        abstractNum1.Append(level4);
        abstractNum1.Append(level5);
        abstractNum1.Append(level6);
        abstractNum1.Append(level7);
        abstractNum1.Append(level8);
        abstractNum1.Append(level9);
        return abstractNum1;
    }

    public AbstractNum GenerateAbstractNum_Decimal_Structured(int abstractNumId)
    {
        AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = abstractNumId };

        Level level1 = new Level() { LevelIndex = 0 };
        StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText1 = new LevelText() { Val = " %1 " };
        LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

        Tabs tabs1 = new Tabs();
        TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 720 };

        tabs1.Append(tabStop1);
        Indentation indentation1 = new Indentation() { Left = "720", Hanging = "360" };

        previousParagraphProperties1.Append(tabs1);
        previousParagraphProperties1.Append(indentation1);

        level1.Append(startNumberingValue1);
        level1.Append(numberingFormat1);
        level1.Append(levelText1);
        level1.Append(levelJustification1);
        level1.Append(previousParagraphProperties1);

        Level level2 = new Level() { LevelIndex = 1 };
        StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText2 = new LevelText() { Val = " %1.%2 " };
        LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

        Tabs tabs2 = new Tabs();
        TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 1080 };

        tabs2.Append(tabStop2);
        Indentation indentation2 = new Indentation() { Left = "1080", Hanging = "360" };

        previousParagraphProperties2.Append(tabs2);
        previousParagraphProperties2.Append(indentation2);

        level2.Append(startNumberingValue2);
        level2.Append(numberingFormat2);
        level2.Append(levelText2);
        level2.Append(levelJustification2);
        level2.Append(previousParagraphProperties2);

        Level level3 = new Level() { LevelIndex = 2 };
        StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText3 = new LevelText() { Val = " %1.%2.%3 " };
        LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

        Tabs tabs3 = new Tabs();
        TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 1440 };

        tabs3.Append(tabStop3);
        Indentation indentation3 = new Indentation() { Left = "1440", Hanging = "360" };

        previousParagraphProperties3.Append(tabs3);
        previousParagraphProperties3.Append(indentation3);

        level3.Append(startNumberingValue3);
        level3.Append(numberingFormat3);
        level3.Append(levelText3);
        level3.Append(levelJustification3);
        level3.Append(previousParagraphProperties3);

        Level level4 = new Level() { LevelIndex = 3 };
        StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText4 = new LevelText() { Val = " %1.%2.%3.%4 " };
        LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

        Tabs tabs4 = new Tabs();
        TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 1800 };

        tabs4.Append(tabStop4);
        Indentation indentation4 = new Indentation() { Left = "1800", Hanging = "360" };

        previousParagraphProperties4.Append(tabs4);
        previousParagraphProperties4.Append(indentation4);

        level4.Append(startNumberingValue4);
        level4.Append(numberingFormat4);
        level4.Append(levelText4);
        level4.Append(levelJustification4);
        level4.Append(previousParagraphProperties4);

        Level level5 = new Level() { LevelIndex = 4 };
        StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText5 = new LevelText() { Val = " %1.%2.%3.%4.%5 " };
        LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

        Tabs tabs5 = new Tabs();
        TabStop tabStop5 = new TabStop() { Val = TabStopValues.Number, Position = 2160 };

        tabs5.Append(tabStop5);
        Indentation indentation5 = new Indentation() { Left = "2160", Hanging = "360" };

        previousParagraphProperties5.Append(tabs5);
        previousParagraphProperties5.Append(indentation5);

        level5.Append(startNumberingValue5);
        level5.Append(numberingFormat5);
        level5.Append(levelText5);
        level5.Append(levelJustification5);
        level5.Append(previousParagraphProperties5);

        Level level6 = new Level() { LevelIndex = 5 };
        StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText6 = new LevelText() { Val = " %1.%2.%3.%4.%5.%6 " };
        LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

        Tabs tabs6 = new Tabs();
        TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 2520 };

        tabs6.Append(tabStop6);
        Indentation indentation6 = new Indentation() { Left = "2520", Hanging = "360" };

        previousParagraphProperties6.Append(tabs6);
        previousParagraphProperties6.Append(indentation6);

        level6.Append(startNumberingValue6);
        level6.Append(numberingFormat6);
        level6.Append(levelText6);
        level6.Append(levelJustification6);
        level6.Append(previousParagraphProperties6);

        Level level7 = new Level() { LevelIndex = 6 };
        StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText7 = new LevelText() { Val = " %1.%2.%3.%4.%5.%6.%7 " };
        LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

        Tabs tabs7 = new Tabs();
        TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

        tabs7.Append(tabStop7);
        Indentation indentation7 = new Indentation() { Left = "2880", Hanging = "360" };

        previousParagraphProperties7.Append(tabs7);
        previousParagraphProperties7.Append(indentation7);

        level7.Append(startNumberingValue7);
        level7.Append(numberingFormat7);
        level7.Append(levelText7);
        level7.Append(levelJustification7);
        level7.Append(previousParagraphProperties7);

        Level level8 = new Level() { LevelIndex = 7 };
        StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText8 = new LevelText() { Val = " %1.%2.%3.%4.%5.%6.%7.%8 " };
        LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

        Tabs tabs8 = new Tabs();
        TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 3240 };

        tabs8.Append(tabStop8);
        Indentation indentation8 = new Indentation() { Left = "3240", Hanging = "360" };

        previousParagraphProperties8.Append(tabs8);
        previousParagraphProperties8.Append(indentation8);

        level8.Append(startNumberingValue8);
        level8.Append(numberingFormat8);
        level8.Append(levelText8);
        level8.Append(levelJustification8);
        level8.Append(previousParagraphProperties8);

        Level level9 = new Level() { LevelIndex = 8 };
        StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
        NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
        LevelText levelText9 = new LevelText() { Val = " %1.%2.%3.%4.%5.%6.%7.%8.%9 " };
        LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

        PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

        Tabs tabs9 = new Tabs();
        TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 3600 };

        tabs9.Append(tabStop9);
        Indentation indentation9 = new Indentation() { Left = "3600", Hanging = "360" };

        previousParagraphProperties9.Append(tabs9);
        previousParagraphProperties9.Append(indentation9);

        level9.Append(startNumberingValue9);
        level9.Append(numberingFormat9);
        level9.Append(levelText9);
        level9.Append(levelJustification9);
        level9.Append(previousParagraphProperties9);

        abstractNum1.Append(level1);
        abstractNum1.Append(level2);
        abstractNum1.Append(level3);
        abstractNum1.Append(level4);
        abstractNum1.Append(level5);
        abstractNum1.Append(level6);
        abstractNum1.Append(level7);
        abstractNum1.Append(level8);
        abstractNum1.Append(level9);
        return abstractNum1;
    }

}