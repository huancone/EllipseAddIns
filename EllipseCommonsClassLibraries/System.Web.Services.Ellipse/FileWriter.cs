using System;
using System.IO;

namespace System.Web.Services.Ellipse
{
    public static class FileWriter
    {
        public static string NormalizePath(string path)
        {
            return Path.GetFullPath(new Uri(path).LocalPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                .ToUpperInvariant();
        }

        public static string NormalizePath(string path, bool expandEnvironment)
        {
            return NormalizePath(expandEnvironment ? Environment.ExpandEnvironmentVariables(path) : path);
        }

        public static void WriteTextToFile(string text, string filename, string urlPath = "")
        {
            //if (!string.IsNullOrWhiteSpace(urlPath) &&
            //    !(urlPath.EndsWith("" + Path.DirectorySeparatorChar) || urlPath.EndsWith("" + Path.AltDirectorySeparatorChar)))
            //    urlPath = urlPath + Path.DirectorySeparatorChar;

            if (urlPath == null)
                urlPath = "";
            File.WriteAllText(Path.Combine(urlPath, filename), text);
        }

        public static void WriteTextToFile(string[] text, string filename, string urlPath = "")
        {
            //if (!string.IsNullOrWhiteSpace(urlPath) &&
            //    !(urlPath.EndsWith("" + Path.DirectorySeparatorChar) ||
            //      urlPath.EndsWith("" + Path.AltDirectorySeparatorChar)))
            //    urlPath = urlPath + Path.DirectorySeparatorChar;

            if (urlPath == null)
                urlPath = "";
            File.WriteAllLines(Path.Combine(urlPath, filename), text);
        }

        public static void AppendTextToFile(string text, string filename, string urlPath = "")
        {
            //if (!string.IsNullOrWhiteSpace(urlPath) &&
            //    !(urlPath.EndsWith("" + Path.DirectorySeparatorChar) ||
            //      urlPath.EndsWith("" + Path.AltDirectorySeparatorChar)))
            //    urlPath = urlPath + Path.DirectorySeparatorChar;

            if (urlPath == null)
                urlPath = "";

            using (var file = new StreamWriter(Path.Combine(urlPath, filename), true))
            {
                file.WriteLine(text);
                file.Flush();
                file.Close();
            }
        }

        public static void CreateDirectory(string directoryPath)
        {
            // Determine whether the directory exists.
            if (Directory.Exists(directoryPath))
                return;

            // Try to create the directory.
            Directory.CreateDirectory(directoryPath);
        }

        public static void DeleteDirectory(string directoryPath)
        {
            // Determine whether the directory exists.
            if (!Directory.Exists(directoryPath))
                return;

            // Try to delete the directory.
            var di = new DirectoryInfo(directoryPath);
            di.Delete();
        }

        public static void DeleteFile(string directoryPath, string fileName)
        {
            DeleteFile(Path.Combine(directoryPath, fileName));
        }

        public static void DeleteFile(string urlFileName)
        {
            // Determine whether the file exists.
            if (!File.Exists(urlFileName))
                return;

            // Try to delete the file.
            var fi = new FileInfo(urlFileName);
            fi.Delete();
        }

        public static bool CheckDirectoryExist(string directoryPath)
        {
            // Determine whether the directory exists.
            return Directory.Exists(directoryPath);
        }

        public static void CopyFileToDirectory(string fileName, string sourcePath, string targetPath,
            bool overwrite = true)
        {
            var sourceFile = Path.Combine(sourcePath, fileName);
            var destFile = Path.Combine(targetPath, fileName);

            File.Copy(sourceFile, destFile, overwrite);
        }

        public static void MoveFileToDirectory(string sourceFileName, string sourcePath, string targetFileName,
            string targetPath)
        {
            var sourceFile = Path.Combine(sourcePath, sourceFileName);
            var destFile = Path.Combine(targetPath, targetFileName);

            File.Move(sourceFile, destFile);
        }
    }
}