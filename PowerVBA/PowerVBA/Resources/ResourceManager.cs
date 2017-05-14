using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace PowerVBA.Resources
{
    public static class ResourceManager
    {
        #region [ Resource ]
        public static Stream GetStreamResource(string path, bool GetPath = true)
        {
            string uri = path;
            if (GetPath) uri = BuildResourceUri(path);


            return Assembly.GetExecutingAssembly()
                .GetManifestResourceStream(uri);
        }

        public static (string, string)[] GetCodeDatas(string FolderPath)
        {
            var Codes = new List<(string, string)>();

            foreach (string path in GetCodePaths(FolderPath)) Codes.Add((path, GetTextResource(path, false)));

            return Codes.ToArray();
        }

        public static string[] GetCodePaths(string FolderPath)
        {
            string uri = BuildResourceUri(FolderPath);

            var paths = Assembly.GetExecutingAssembly().GetManifestResourceNames().Where((i) => i.StartsWith(uri));

            return paths.ToArray();
        }

        public static string GetTextResource(string path, bool GetPath = true)
        {
            Stream s = GetStreamResource(path, GetPath);

            using (var sr = new StreamReader(s)) return sr.ReadToEnd();
        }

        private static string BuildResourceUri(string path)
        {
            return $"PowerVBA.Resources.{path.Replace('/', '.')}";
        }
        #endregion
    }
}
