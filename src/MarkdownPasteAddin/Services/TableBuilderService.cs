using MarkdownPasteAddin.Models;
using Word = Microsoft.Office.Interop.Word;

namespace MarkdownPasteAddin.Services;

public class TableBuilderService
{
    public Word.Table CreateTable(Word.Range range, TableChunk chunk)
    {
        int totalRows = chunk.RowCount + 1; // +1 for header
        int totalCols = chunk.ColumnCount;

        var table = range.Tables.Add(range, totalRows, totalCols);

        // Apply built-in grid style
        table.set_Style("Table Grid");

        // Fill header row
        for (int col = 0; col < totalCols; col++)
        {
            var cell = table.Cell(1, col + 1);
            cell.Range.Text = chunk.Headers[col];
            ApplyCellFormatting(cell, isHeader: true);
        }

        // Fill data rows
        for (int row = 0; row < chunk.Rows.Count; row++)
        {
            for (int col = 0; col < totalCols && col < chunk.Rows[row].Count; col++)
            {
                var cell = table.Cell(row + 2, col + 1);
                cell.Range.Text = chunk.Rows[row][col];
                ApplyCellFormatting(cell, isHeader: false);
            }
        }

        // Set column alignments
        for (int col = 0; col < totalCols && col < chunk.Alignments.Count; col++)
        {
            var alignment = chunk.Alignments[col] switch
            {
                ColumnAlignment.Center => Word.WdParagraphAlignment.wdAlignParagraphCenter,
                ColumnAlignment.Right => Word.WdParagraphAlignment.wdAlignParagraphRight,
                _ => Word.WdParagraphAlignment.wdAlignParagraphLeft
            };

            for (int row = 0; row < totalRows; row++)
            {
                table.Cell(row + 1, col + 1).Range.ParagraphFormat.Alignment = alignment;
            }
        }

        // Auto-fit to window
        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

        return table;
    }

    private void ApplyCellFormatting(Word.Cell cell, bool isHeader)
    {
        if (isHeader)
        {
            cell.Range.Font.Bold = 1;
            cell.Range.Shading.BackgroundPatternColor =
                Word.WdColor.wdColorGray10;
        }
    }
}
