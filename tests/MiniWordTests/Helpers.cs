using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MiniWordTests
{
    internal static class Helpers
    {
        internal static string GetZipFileContent(string zipPath, string filePath)
        {
            var ns = new XmlNamespaceManager(new NameTable());
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            using (var stream = File.OpenRead(zipPath))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
            {
                var entry = archive.Entries.Single(w => w.FullName == filePath);
                using (var sheetStream = entry.Open())
                {
                    var doc = XDocument.Load(sheetStream);
                    return doc.ToString();
                }
            }
        }
    }
}
