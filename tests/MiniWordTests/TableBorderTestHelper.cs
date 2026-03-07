using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MiniWordTests
{
    internal static class TableBorderTestHelper
    {
        /// <summary>
        /// Creates a minimal Word template that contains a table with:
        /// <list type="bullet">
        ///   <item>A header row with static text.</item>
        ///   <item>A data row with <c>{{Items.Name}}</c> and <c>{{Items.Value}}</c> placeholders.</item>
        ///   <item>A dedicated border-control row with a <c>{{TableBorder}}</c> placeholder.</item>
        /// </list>
        /// </summary>
        internal static void CreateTableBorderTemplate(string filePath)
        {
            using var doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Table(
                        new TableRow(
                            new TableCell(new Paragraph(new Run(new Text("Name")))),
                            new TableCell(new Paragraph(new Run(new Text("Value")))),
                            new TableCell(new Paragraph(new Run(new Text("{{TableBorder}}"))))
                        ),
                        new TableRow(
                            new TableCell(new Paragraph(new Run(new Text("{{Items.Name}}")))),
                            new TableCell(new Paragraph(new Run(new Text("{{Items.Value}}")))),
                            new TableCell(new Paragraph(new Run(new Text(""))))
                        )
                    )
                )
            );
            doc.Save();
        }
    }
}
