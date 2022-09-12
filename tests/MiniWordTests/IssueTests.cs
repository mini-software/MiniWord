using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace MiniWordTests
{
    public class IssueTests
    {
        /// <summary>
        /// [Paragraph replace by tag · Issue #4 · mini-software/MiniWord](https://github.com/mini-software/MiniWord/issues/4)
        /// </summary>
        [Fact]
        public void TestIssue4()
        {
			var path = PathHelper.GetTempPath();
            var templatePath = PathHelper.GetFile("TestBasicFill.docx");
			var value = new Dictionary<string, object>()
			{
                ["Company_Name"] = "MiniSofteware",
                ["Name"] = "Jack",
				["CreateDate"] = new DateTime(2021, 01, 01),
				["VIP"] = true,
				["Points"] = 123,
				["APP"] = "Demo APP",
			};
			MiniSoftware.MiniWord.SaveAsByTemplate(path, templatePath, value);
			Console.WriteLine(path);
		}
    }

    internal static class PathHelper
    {
        public static string GetFile(string fileName,string folderName="docx")
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
