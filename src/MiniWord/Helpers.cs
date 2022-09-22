namespace MiniSoftware
{
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Dynamic;
    using System.IO;

    internal static partial class Helpers
	{
		public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
	}
}