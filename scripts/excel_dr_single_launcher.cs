using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;

internal static class ExcelDrSingleLauncher
{
    private const string ResourceName = "ExcelDrPayload.zip";

    [STAThread]
    private static int Main()
    {
        try
        {
            string packageVersion = GetPackageVersion();
            string root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Excel-Dr",
                "single-exe",
                packageVersion);
            string appExe = Path.Combine(root, "Excel-Dr.exe");

            using (Mutex mutex = new Mutex(false, "ExcelDrSingleLauncher_" + packageVersion))
            {
                mutex.WaitOne();
                try
                {
                    if (!IsExtracted(root))
                    {
                        ExtractPayload(root);
                    }
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = appExe,
                WorkingDirectory = root,
                UseShellExecute = true
            });
            return 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Excel-Dr 启动失败：\n\n" + ex.Message,
                "Excel-Dr",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static void ExtractPayload(string targetRoot)
    {
        string parent = Path.GetDirectoryName(targetRoot);
        if (parent == null)
        {
            throw new InvalidOperationException("无法定位解压目录。");
        }

        Directory.CreateDirectory(parent);
        string tempRoot = Path.Combine(parent, "extracting-" + Guid.NewGuid().ToString("N"));
        string zipPath = Path.Combine(tempRoot, "payload.zip");
        Directory.CreateDirectory(tempRoot);

        try
        {
            using (Stream payload = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
            {
                if (payload == null)
                {
                    throw new InvalidOperationException("单文件包缺少内置运行资源。");
                }

                using (FileStream output = File.Create(zipPath))
                {
                    payload.CopyTo(output);
                }
            }

            ZipFile.ExtractToDirectory(zipPath, tempRoot);
            File.Delete(zipPath);

            if (Directory.Exists(targetRoot))
            {
                Directory.Delete(targetRoot, true);
            }

            Directory.Move(tempRoot, targetRoot);
        }
        catch
        {
            if (Directory.Exists(tempRoot))
            {
                Directory.Delete(tempRoot, true);
            }
            throw;
        }
    }

    private static bool IsExtracted(string root)
    {
        return File.Exists(Path.Combine(root, "Excel-Dr.exe"))
            && File.Exists(Path.Combine(root, "hub.dll"))
            && File.Exists(Path.Combine(root, "flutter_windows.dll"))
            && Directory.Exists(Path.Combine(root, "data"));
    }

    private static string GetPackageVersion()
    {
        using (Stream payload = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
        {
            if (payload == null)
            {
                throw new InvalidOperationException("单文件包缺少内置运行资源。");
            }

            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(payload);
                return BitConverter.ToString(hash).Replace("-", "");
            }
        }
    }
}
