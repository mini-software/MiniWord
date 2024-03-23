namespace MiniSoftware
{
    using DocumentFormat.OpenXml.Office2013.Excel;
    using MiniSoftware.Extensions;
    using MiniSoftware.Utility;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Dynamic;
    using System.IO;
    using System.Linq.Expressions;
    using System.Threading;
    using System.Threading.Tasks;

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

        public static async Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value,CancellationToken token = default(CancellationToken))
        {
            await SaveAsByTemplateImplAsync(stream, templateBytes, value.ToDictionary(),token).ConfigureAwait(false);
        }

        public static async Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value,CancellationToken token = default(CancellationToken))
        {
            await SaveAsByTemplateImplAsync(stream, await GetByteAsync(templatePath), value.ToDictionary(),token).ConfigureAwait(false);
        }

        public static async Task SaveAsByTemplateAsync(string path, string templatePath, object value,CancellationToken token = default(CancellationToken))
        {
            using (var stream = FileHelper.CreateAsync(path))
                await SaveAsByTemplateImplAsync(await stream, await GetByteAsync(templatePath), value.ToDictionary(),token);
        }

        public static async Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value,CancellationToken token = default(CancellationToken))
        {
            using (var stream = FileHelper.CreateAsync(path))
                await SaveAsByTemplateImplAsync(await stream, templateBytes, value.ToDictionary(),token);
        }
    }
}