using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("Kinderbetreuung")]
[assembly: AssemblyDescription("Lokale Dokumentation und Auswertung von Kinderbetreuungskosten")]
[assembly: AssemblyCompany("Lorenz Köcke")]
[assembly: AssemblyProduct("Kinderbetreuung")]
[assembly: AssemblyCopyright("Copyright © 2026 Lorenz Köcke")]
[assembly: AssemblyVersion("__VERSION__.0")]
[assembly: AssemblyFileVersion("__VERSION__.0")]

internal static class Program
{
    private const string Version = "__VERSION__";
    private const string ResourceName = "Kinderbetreuung.Payload.zip";

    [STAThread]
    private static void Main()
    {
        try
        {
            string baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kinderbetreuung", "Runtime", Version);

            Directory.CreateDirectory(baseDir);
            ExtractPayload(baseDir);

            string script = Path.Combine(baseDir, "App.ps1");
            if (!File.Exists(script))
                throw new FileNotFoundException("App.ps1 wurde im Programmpaket nicht gefunden.");

            var start = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + script + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = baseDir
            };

            using (var process = Process.Start(start))
            {
                if (process == null)
                    throw new InvalidOperationException("PowerShell konnte nicht gestartet werden.");
                process.WaitForExit();
                if (process.ExitCode != 0)
                    throw new InvalidOperationException("Kinderbetreuung wurde mit Fehlercode " + process.ExitCode + " beendet.");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                "Kinderbetreuung " + Version + " - Startfehler",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static void ExtractPayload(string destination)
    {
        Assembly asm = Assembly.GetExecutingAssembly();
        using (Stream stream = asm.GetManifestResourceStream(ResourceName))
        {
            if (stream == null)
                throw new InvalidOperationException("Interne Programmdateien fehlen.");

            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                foreach (var entry in archive.Entries)
                {
                    string target = Path.Combine(destination, entry.FullName.Replace('/', Path.DirectorySeparatorChar));
                    string fullTarget = Path.GetFullPath(target);
                    string fullBase = Path.GetFullPath(destination) + Path.DirectorySeparatorChar;

                    if (!fullTarget.StartsWith(fullBase, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Ungueltiger Dateipfad im Programmpaket.");

                    if (string.IsNullOrEmpty(entry.Name))
                    {
                        Directory.CreateDirectory(fullTarget);
                        continue;
                    }

                    Directory.CreateDirectory(Path.GetDirectoryName(fullTarget));
                    using (Stream input = entry.Open())
                    using (FileStream output = new FileStream(fullTarget, FileMode.Create, FileAccess.Write, FileShare.None))
                        input.CopyTo(output);
                }
            }
        }
    }
}
