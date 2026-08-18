using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

internal static class Program
{
    private const string Version = "1.0.18";
    private const string ScriptResource = "Kinderbetreuung.EmbeddedApp.ps1";

    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        try
        {
            string script = ReadEmbeddedText(ScriptResource);

            using (var runspace = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace())
            {
                // WinForms-Autocomplete, Dialoge und weitere UI-Funktionen benoetigen STA.
                runspace.ApartmentState = System.Threading.ApartmentState.STA;
                runspace.ThreadOptions = System.Management.Automation.Runspaces.PSThreadOptions.ReuseThread;
                runspace.Open();

                using (PowerShell ps = PowerShell.Create())
                {
                    ps.Runspace = runspace;

                    // Kennzeichnet den eingebetteten Release-Modus.
                    ps.AddScript("$script:EmbeddedMode = $true");
                    ps.Invoke();

                    ps.Commands.Clear();
                    ps.AddScript(script, useLocalScope: false);

                    Collection<PSObject> result = ps.Invoke();

                    if (ps.HadErrors)
                    {
                        StringBuilder errors = new StringBuilder();
                        foreach (var error in ps.Streams.Error)
                            errors.AppendLine(error.ToString());

                        throw new InvalidOperationException(errors.ToString().Trim());
                    }
                }
            }
        }
        catch (Exception ex)
        {
            string dataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kinderbetreuung");
            Directory.CreateDirectory(dataDir);

            string errorFile = Path.Combine(dataDir, "startfehler.txt");
            try
            {
                File.WriteAllText(
                    errorFile,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                    Environment.NewLine +
                    ex.ToString(),
                    Encoding.UTF8);
            }
            catch { }

            MessageBox.Show(
                "Startfehler:" + Environment.NewLine +
                ex.Message + Environment.NewLine + Environment.NewLine +
                "Details:" + Environment.NewLine + errorFile,
                "Kinderbetreuung " + Version,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string ReadEmbeddedText(string resourceName)
    {
        Assembly asm = Assembly.GetExecutingAssembly();
        using (Stream stream = asm.GetManifestResourceStream(resourceName))
        {
            if (stream == null)
                throw new InvalidOperationException(
                    "Die eingebettete Anwendung konnte nicht gefunden werden.");

            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
                return reader.ReadToEnd();
        }
    }
}
