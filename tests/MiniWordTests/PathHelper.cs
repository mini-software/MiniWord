using System;
using System.IO;

namespace MiniWordTests
{
    internal static class PathHelper
    {
        public static string GetFile(string fileName, string folderName = "docx")
        {
            return $@"../../../../../samples/{folderName}/{fileName}";
        }

        public static string GetTempPath(string extension = "docx")
        {
            var method = (new System.Diagnostics.StackTrace()).GetFrame(1).GetMethod();

            var path = Path.Combine(Path.GetTempPath(), $"{method.DeclaringType.Name}_{method.Name}.{extension}").Replace("<", string.Empty).Replace(">", string.Empty);
            if (File.Exists(path))
                File.Delete(path);
            return path;
        }

        public static string GetTempFilePath(string extension = "docx")
        {
            return Path.GetTempPath() + Guid.NewGuid().ToString() + "." + extension;
        }
    }
}
