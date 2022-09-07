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

    public enum TableCellBorderSelection { Left, Top, Right, Bottom, All };

    public static partial class DocXExt
    {

        /// <summary>
        /// mimic of margin using indentation(start:left, end:right) and spacing(before:top, after:bottom)
        /// on cell paragraphs
        /// </summary>
        /// <param name="cell">table cell</param>
        /// <param name="leftMM">left margin (mm)</param>
        /// <param name="topMM">top margin (mm)</param>
        /// <param name="rightMM">right margin (mm)</param>
        /// <param name="bottomMM">bottom margin (mm)</param>
        /// <returns>table cell</returns>
        public static TableCell SetPadding(this TableCell cell,
            double? leftMM = null, double? topMM = null, double? rightMM = null, double? bottomMM = null)
        {
            var paragraphs = cell.Elements<Paragraph>().ToList();

            for (int i = 0; i < paragraphs.Count; ++i)
            {
                var paragraph = paragraphs[i];

                paragraph.SetMargin(leftMM, i == 0 ? topMM : null, rightMM, i == paragraphs.Count - 1 ? bottomMM : null);
            }

            return cell;
        }

        /// <summary>
        /// mimic of margin using indentation(start:left, end:right) and spacing(before:top, after:bottom)
        /// on cell paragraphs
        /// </summary>
        /// <param name="cell">table cell</param>
        /// <param name="paddingMM">padding (mm)</param>        
        /// <returns>table cell</returns>
        public static TableCell SetPadding(this TableCell cell, double? paddingMM = null) =>
            cell.SetPadding(paddingMM, paddingMM, paddingMM, paddingMM);

        public static Paragraph SetParagraph(this TableCell cell, string? text = null, Action<Run>? action = null)
        {
            var toremove = cell.Elements<Paragraph>().ToList();
            foreach (var x in toremove) x.Remove();

            var newParagraph = new Paragraph();
            cell.Append(newParagraph);

            if (text is not null)
            {
                var newRun = new Run();
                newParagraph.Append(newRun);
                {
                    var newText = new Text(text);
                    newRun.Append(newText);
                    if (action is not null) action(newRun);
                }
            }

            return newParagraph;
        }

        /// <summary>
        /// add columns in bunch
        /// </summary>
        /// <param name="table">table which adds columns</param>
        /// <param name="columnCount">number of columns to add</param>
        /// <param name="colWidthMM">initial column width</param>
        /// <param name="action">action to perform on added column</param>
        /// <returns>table</returns>
        public static Table AddColumns(this Table table, int columnCount, double colWidthMM, Action<GridColumn>? action = null)
        {
            for (int i = 0; i < columnCount; ++i) table.AddColumn(colWidthMM, action);

            return table;
        }

        /// <summary>
        /// adds a column to the table
        /// </summary>
        /// <param name="table">table</param>
        /// <param name="colWidthMM">column width (mm) ; note: if table is in % then column widths will normalized</param>
        /// <param name="action">(optional) action on created GridColumn</param>
        /// <returns>grid column</returns>
        public static Table AddColumn(this Table table, double colWidthMM, Action<GridColumn>? action = null)
        {
            var grid = table.Grid();

            var gridColumn = new GridColumn { Width = colWidthMM.MMToTwip().ToString(CultureInfo.InvariantCulture) };
            grid.Append(gridColumn);

            foreach (var row in table.GetRows())
            {
                var tableCell = new TableCell();
                row.Append(tableCell);
            }

            action?.Invoke(gridColumn);

            return table;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static Table GetTable(this TableRow tableRow) => (Table)tableRow.Parent!;

        /// <summary>
        /// add a row to the table initializing cells settings default empty paragraph
        /// </summary>
        /// <param name="table">table</param>
        /// <param name="action">action on row</param>        
        /// <returns>table row</returns>
        public static Table AddRow(this Table table, Action<TableRow>? action = null)
        {
            var tableRow = new TableRow();

            table.Append(tableRow);

            var grid = table.Grid();

            int colCount = table.GetColumnCount();

            for (int cidx = 0; cidx < colCount; ++cidx)
            {
                var tableCell = new TableCell();

                tableRow.Append(tableCell);

                tableCell.AddParagraph();
            }

            action?.Invoke(tableRow);

            return table;
        }

        static void SetTableBorderOutside(
            TableRow row, TableCell cell,
            int rowCount, int colCount,
            int rowIdx, int colIdx,
            ref LeftBorder? leftBorder,
            ref TopBorder? topBorder,
            ref RightBorder? rightBorder,
            ref BottomBorder? bottomBorder,
            Action<BorderType>? customAction)
        {
            if (colIdx == 0)
            {
                leftBorder = new LeftBorder();
                customAction?.Invoke(leftBorder);
            }
            if (rowIdx == 0)
            {
                topBorder = new TopBorder();
                customAction?.Invoke(topBorder);
            }
            if (colIdx == colCount - 1)
            {
                rightBorder = new RightBorder();
                customAction?.Invoke(rightBorder);
            }
            if (rowIdx == rowCount - 1)
            {
                bottomBorder = new BottomBorder();
                customAction?.Invoke(bottomBorder);
            }
        }

        /// <summary>
        /// set table borders for only outside borders mode
        /// </summary>
        /// <param name="table">table</param>
        /// <param name="type">type of border</param>
        public static Table SetBorderOutside(this Table table, BorderValues type = BorderValues.Single) =>
            table.SetBorder(SetTableBorderOutside, borderType => borderType.Val = type);

        static void SetTableBorderAll(
            TableRow row, TableCell cell,
            int rowCount, int colCount,
            int rowIdx, int colIdx,
            ref LeftBorder? leftBorder,
            ref TopBorder? topBorder,
            ref RightBorder? rightBorder,
            ref BottomBorder? bottomBorder,
            Action<BorderType>? customAction)
        {
            leftBorder = new LeftBorder();
            customAction?.Invoke(leftBorder);

            topBorder = new TopBorder();
            customAction?.Invoke(topBorder);

            rightBorder = new RightBorder();
            customAction?.Invoke(rightBorder);

            bottomBorder = new BottomBorder();
            customAction?.Invoke(bottomBorder);
        }

        /// <summary>
        /// set table borders all cell borders
        /// </summary>
        /// <param name="table">table</param>
        /// <param name="type">type of border</param>
        public static Table SetBorderAll(this Table table, BorderValues type = BorderValues.Single) =>
            table.SetBorder(SetTableBorderAll, borderType => borderType.Val = type);

        /// <summary>
        /// delegate to set cell border
        /// </summary>
        /// <param name="row">current row</param>
        /// <param name="cell">current cell</param>        
        /// <param name="rowCount">numbers of rows</param>
        /// <param name="colCount">numbers of columns</param>
        /// <param name="rowIdx">current row zerobased index</param>
        /// <param name="colIdx">current col zerobased index</param>        
        /// <param name="leftBorder">left border reference ( if null mean it was not exists, can be reassigned to a new )</param>
        /// <param name="topBorder">top border reference ( if null mean it was not exists, can be reassigned to a new )</param>
        /// <param name="rightBorder">right border reference ( if null mean it was not exists, can be reassigned to a new )</param>
        /// <param name="bottomBorder">bottom border reference ( if null mean it was not exists, can be reassigned to a new )</param>
        /// <param name="customAction">custom action flowing from SetBorder method caller</param>
        public delegate void SetTableBorderTypeDelegate(
            TableRow row, TableCell cell,
            int rowCount, int colCount,
            int rowIdx, int colIdx,
            ref LeftBorder? leftBorder,
            ref TopBorder? topBorder,
            ref RightBorder? rightBorder,
            ref BottomBorder? bottomBorder,
            Action<BorderType>? customAction);

        /// <summary>
        /// set border of table ; custom border foreach cell can specified using action
        /// </summary>
        /// <param name="table">table</param>        
        /// <param name="action">custom action foreach cell</param>
        /// <param name="customAction">custom action that will applied to border changed by main action</param>
        /// <returns>table</returns>
        public static Table SetBorder(this Table table, SetTableBorderTypeDelegate action, Action<BorderType>? customAction = null)
        {
            var columnCount = table.GetColumnCount();
            var rowsCount = table.GetRowCount();

            for (int r = 0; r < rowsCount; ++r)
            {
                var row = table.GetRow(r);

                for (int c = 0; c < columnCount; ++c)
                {
                    var cell = row.GetCell(c);

                    var borders = cell
                        .GetOrCreate<TableCellProperties>(createIfNotExists: false)?
                        .GetOrCreate<TableCellBorders>(createIfNotExists: false);

                    var leftBorder = borders is null ? null : borders.LeftBorder;
                    var topBorder = borders is null ? null : borders.TopBorder;
                    var rightBorder = borders is null ? null : borders.RightBorder;
                    var bottomBorder = borders is null ? null : borders.BottomBorder;

                    action.Invoke(
                        row, cell,
                        rowsCount, columnCount,
                        r, c,
                        ref leftBorder, ref topBorder, ref rightBorder, ref bottomBorder, customAction);

                    if (borders is null)
                    {
                        var leftApplied = leftBorder is not null;
                        var topApplied = topBorder is not null;
                        var rightApplied = rightBorder is not null;
                        var bottomApplied = bottomBorder is not null;

                        if (leftApplied || topApplied || rightApplied || bottomApplied)
                        {
                            borders = cell.GetTableCellProperties().GetTableCellBorders();

                            if (leftApplied) borders.LeftBorder = leftBorder;
                            if (topApplied) borders.TopBorder = topBorder;
                            if (rightApplied) borders.RightBorder = rightBorder;
                            if (bottomApplied) borders.BottomBorder = bottomBorder;
                        }
                    }
                }
            }

            return table;
        }

        /// <summary>
        /// retrieve table rows
        /// </summary>
        /// <param name="table">table</param>        
        /// <returns>enumerable of TableRow</returns>
        public static IEnumerable<TableRow> GetRows(this Table table) => table.Elements<TableRow>();

        /// <summary>
        /// retrieve table row zerobased index
        /// </summary>
        /// <param name="table">table</param>
        /// <param name="row">row (0 is first)</param>        
        /// <returns>table row at given zero based index</returns>
        public static TableRow GetRow(this Table table, int row) => table.Elements<TableRow>().Skip(row).First();

        /// <summary>
        /// retrieve table cell zerobased index from given row;
        /// note: to set paragraph use SetParagraph() instead of AddParagraph() if want to set first one
        /// because initial cell have a dummy paragraph.
        /// </summary>
        /// <param name="row">row from which retrieve cell</param>
        /// <param name="col">zero-based index of cell column</param>        
        /// <returns>table row column at given zero based index</returns>
        public static TableCell GetCell(this TableRow row, int col) => row.Elements<TableCell>().Skip(col).First();

        static TableCellProperties GetTableCellProperties(this TableCell cell) =>
            cell.GetOrCreate<TableCellProperties>(createIfNotExists: true, insertAtIdx: 0)!;

        static TableCellBorders GetTableCellBorders(this TableCellProperties tableCellProperties) =>
            tableCellProperties.GetOrCreate<TableCellBorders>(createIfNotExists: true)!;

        public static BorderType SetType(this BorderType borderType,
            BorderValues type = BorderValues.Single)
        {
            borderType.Val = type;

            return borderType;
        }

        /// <summary>
        /// set a specific border ( left, top, right, bottom ) or all at once
        /// </summary>
        /// <param name="cell">cell</param>
        /// <param name="action">action to apply to border selection</param>
        /// <param name="borderSelectionType">border selection type</param>
        /// <returns>cell</returns>
        public static TableCell SetBorder(this TableCell cell,
            Action<BorderType> action,
            TableCellBorderSelection borderSelectionType = TableCellBorderSelection.All)
        {
            var tableCellBorders = cell
                .GetTableCellProperties()
                .GetTableCellBorders();

            void SetLeft()
            {
                var brd = new LeftBorder();
                tableCellBorders.LeftBorder = brd;
                action(brd);
            }

            void SetTop()
            {
                var brd = new TopBorder();
                tableCellBorders.TopBorder = brd;
                action(brd);
            }

            void SetRight()
            {
                var brd = new RightBorder();
                tableCellBorders.RightBorder = brd;
                action(brd);
            }

            void SetBottom()
            {
                var brd = new BottomBorder();
                tableCellBorders.BottomBorder = brd;
                action(brd);
            }

            switch (borderSelectionType)
            {
                case TableCellBorderSelection.Left: SetLeft(); break;
                case TableCellBorderSelection.Top: SetTop(); break;
                case TableCellBorderSelection.Right: SetRight(); break;
                case TableCellBorderSelection.Bottom: SetBottom(); break;

                case TableCellBorderSelection.All:
                    SetLeft();
                    SetTop();
                    SetRight();
                    SetBottom();
                    break;
            }

            return cell;
        }

        /// <summary>
        /// retrieve table cell zerobased index from given table
        /// </summary>
        /// <param name="table">table</param>
        /// <param name="row">zero-based index of row</param>
        /// <param name="col">zero-based index of column</param>
        /// <returns>cell at given zero-based index of row,col</returns>
        public static TableCell GetCell(this Table table, int row, int col) => table.GetRow(row).GetCell(col);

        /// <summary>
        /// retrieve the number of table columns
        /// </summary>
        /// <param name="table">table</param>        
        /// <returns>number of columns</returns>
        public static int GetColumnCount(this Table table) => table.Grid().Elements<GridColumn>().Count();

        /// <summary>
        /// retrieve the number of table rows
        /// </summary>
        /// <param name="table">table</param>        
        /// <returns>number of rows</returns>
        public static int GetRowCount(this Table table) => table.Elements<TableRow>().Count();

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static TableGrid Grid(this Table table) => table.GetOrCreate<TableGrid>(createIfNotExists: true)!;

        /// <summary>
        /// add table
        /// </summary>
        /// <param name="paragraphBefore">paragraph before the table</param>
        /// <param name="tableWidthPercent">table width percent (0..100) or null for auto</param>
        /// <param name="align">table alignment</param>
        /// <returns>table</returns>
        public static Table AddTable(this Paragraph paragraphBefore,
            double? tableWidthPercent = 100,
            TableRowAlignmentValues align = TableRowAlignmentValues.Left)
        {
            var table = new Table();

            var body = paragraphBefore.GetBody();

            var currentSectionWidth = paragraphBefore.GetCurrentSectionWidthMM();

            {
                var tableProperties1 = new TableProperties();
                TableWidth tableWidth1;
                if (tableWidthPercent is not null)
                    tableWidth1 = new TableWidth
                    {
                        Width = (tableWidthPercent / 100).Value.FactorToPct().ToString(CultureInfo.InvariantCulture),
                        Type = TableWidthUnitValues.Pct
                    };
                else
                    tableWidth1 = new TableWidth
                    {
                        Type = TableWidthUnitValues.Auto
                    };
                var tableJustification1 = new TableJustification { Val = align };
                var tableIndentation1 = new TableIndentation { Width = 0, Type = TableWidthUnitValues.Dxa };
                var tableLayout1 = new TableLayout { Type = TableLayoutValues.Fixed };

                var tableCellMarginDefault1 = new TableCellMarginDefault();
                var topMargin1 = new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa };
                var tableCellLeftMargin1 = new TableCellLeftMargin { Width = 0, Type = TableWidthValues.Dxa };
                var bottomMargin1 = new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa };
                var tableCellRightMargin1 = new TableCellRightMargin { Width = 0, Type = TableWidthValues.Dxa };

                tableCellMarginDefault1.Append(topMargin1);
                tableCellMarginDefault1.Append(tableCellLeftMargin1);
                tableCellMarginDefault1.Append(bottomMargin1);
                tableCellMarginDefault1.Append(tableCellRightMargin1);

                tableProperties1.Append(tableWidth1);
                tableProperties1.Append(tableJustification1);
                tableProperties1.Append(tableIndentation1);
                tableProperties1.Append(tableLayout1);
                tableProperties1.Append(tableCellMarginDefault1);

                table.Append(tableProperties1);
            }

            var tableGrid = new TableGrid();

            table.Append(tableGrid);

            paragraphBefore.InsertAfterSelf(table);

            return table;
        }

        /// <summary>
        /// set shading over table cell extensions
        /// </summary>
        /// <param name="tableCell">table cell which apply shading</param>
        /// <param name="color">shading color</param>
        /// <param name="pattern">shading pattern type</param>
        /// <returns>paragraph</returns>
        public static TableCell SetShading(this TableCell tableCell,
            System.Drawing.Color? color = null,
            ShadingPatternValues pattern = ShadingPatternValues.Clear) =>
            tableCell.SetShading<TableCell, TableCellProperties>(color, pattern);

    }

}