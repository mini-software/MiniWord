namespace MiniSoftware
{
    using System.IO;

    internal static partial class Helpers
	{
		public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
	}
}