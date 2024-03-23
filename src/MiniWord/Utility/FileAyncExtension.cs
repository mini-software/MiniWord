namespace MiniSoftware.Utility
{
    using System.IO;
    using System.Threading.Tasks;
    internal static class FileHelper{
        internal static async Task<FileStream> CreateAsync(string path)
        {
            using (var stream = File.Create(path))
            {
                return await Task.FromResult(stream);
            }
        }
    }
}