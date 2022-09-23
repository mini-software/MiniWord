namespace MiniSoftware
{
    using MiniSoftware.Extensions;
    using System.IO;

    public static partial class MiniWord
    {
        public static void SaveAsByTemplate(string path, string templatePath, object value)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templatePath, value);
        }

        public static void SaveAsByTemplate(string path, byte[] templateBytes, object value)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templateBytes, value);
        }

        public static void SaveAsByTemplate(this Stream stream, string templatePath, object value)
        {
            SaveAsByTemplateImpl(stream, GetBytes(templatePath), value.ToDictionary());
        }

        public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value)
        {
            SaveAsByTemplateImpl(stream, templateBytes, value.ToDictionary());
        }
    }
}