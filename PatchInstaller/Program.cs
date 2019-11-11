using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Text.RegularExpressions;
using System.Threading;
//using System.Management.Automation; //Powershell
using System.Windows.Forms;
//using System.Management.Automation.Runspaces; //Powershell
//using System.Collections.ObjectModel; //Powershell

namespace PatchInstaller
{
    static class Program
    {
        public static string extraction_folder_name;
        public static Form1 form;
        public static DirectoryInfo d;
        public static string[] applicableImageVersions = new string[] { "EVNE2.65P", "EVNE2.6SP", "EVNE2.7SP" };
        public static string[] applicableDRVersions = new string[] { "Canon v2.16", "Canon v2.12", "Canon v2.14" };

        [STAThread]
        static void Main(string[] args)
        {
            #region "Check if application already running and Exit"
            if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                MessageBox.Show("Only one instance of this program is allowed.");
                Environment.Exit(-1);
            }
            #endregion
            #region "Get imageVer and drVersion"
            // Set the three labels: label4, label5, label6
            string imageVer = "";
            string drVersion = "";
            string imageVerPath = @"C:\Shimadzu\Version\ImageVer.txt";
            string[] values;
            using (StreamReader reader = new StreamReader(imageVerPath))
            {
                values = Regex.Split(reader.ReadToEnd(), @"\r\n");
            }
            for (int i = 0; i < values.Count(); i++)
            {
                switch (values[i])
                {
                    case "[ImageVer]":
                        imageVer = values[++i];
                        break;
                }
            }
            if (imageVer.Contains("KM"))
            {
                try
                {
                    string[] lines = File.ReadAllLines(@"C:\Konicaminolta\Console\Data\ReleaseNotes.txt");
                    string version = "0.00";
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("Version Number:"))
                        {
                            version = line.Replace("Version Number: V", "");
                            int end = version.IndexOf('_');
                            version = version.Remove(end, version.Length - end);

                        }
                    }
                    drVersion = "Konica v" + version.ToString();
                }
                catch (Exception e) { }
            }
            else if (imageVer.Contains("NE"))
            {
                try
                {
                    object objec = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Canon Inc\CXDI Controller", "DisplayVersion", null);
                    if (objec != null)
                    {
                        drVersion = "Canon v" + objec.ToString();
                    }
                }
                catch (Exception e) { }
            }
            #endregion
            #region "imageVer check"
            bool included = false;
            foreach (string s in applicableImageVersions)
            {
                if (s == imageVer)
                {
                    included = true;
                }
            }
            if (!included)
            {
                Environment.Exit(-2);
                Application.Exit();
            }
            #endregion
            #region "drVersion check"
            included = false;
            foreach (string s in applicableDRVersions)
            {
                if (drVersion.Contains(s))
                {
                    included = true;
                }
            }
            if (!included)
            {
                Environment.Exit(-3);
                Application.Exit();
            }
            #endregion
            #region "Delete old %temp%\3zx.xxxxxx\ directories"
            try
            {
                string[] files = Directory.GetDirectories(Path.GetTempPath(), "*.*");
                foreach (string f in files)
                {
                    if (f.StartsWith(Path.GetTempPath() + "3z"))
                    {
                        string[] filess = Directory.GetFiles(f, "*.*");
                        foreach (string ff in filess)
                        {
                            File.Delete(ff);
                        }
                        Directory.Delete(f);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Environment.Exit(-4);
            }
            #endregion
            #region "Always run "OnProcessExit" on Application Exit"
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit);
            #endregion
            #region "Create new temp directory for extracting embedded files"
            Random random = new Random((int)DateTime.Now.Ticks);
            extraction_folder_name = "3z" + random.NextDouble().ToString();
            d = Directory.CreateDirectory(Path.GetTempPath() + extraction_folder_name);
            #endregion
            #region "Create the new form"
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            form = new Form1();
            #endregion
            #region "Start the worker thread"
            Thread t;
            if (args[0] == "/install") {
                t = new Thread(install);
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
            }
            else if (args[0] == "/uninstall")
            {
                t = new Thread(uninstall);
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
            } 
            #endregion
            #region "Run the GUI form"
            Application.Run(form);
            #endregion
        }

        private static void install()
        {
            try
            {
                #region "Extract all the Embedded Resources to %temp%\3z[random]\"
                // Create a list of files to extract
                List<string> files_1 = new List<string>();
                string[] resourceNames = form.GetType().Assembly.GetManifestResourceNames();
                foreach (string resourceName in resourceNames)
                {
                    files_1.Add(resourceName);
                }
                files_1.RemoveAt(1);
                files_1.RemoveAt(0);
                form.Invoke((MethodInvoker)delegate {
                    form.progressBar1.Maximum = files_1.Count;
                });

                // Extract the files
                foreach (string file in files_1)
                {
                    ExtractEmbeddedResource(Path.GetTempPath() + extraction_folder_name, "", file);
                }
                #endregion
                #region "Update the RichTextBox"
                form.Invoke((MethodInvoker)delegate {
                    form.richTextBox1.AppendText("\nApplying Patch... ");
                });
                #endregion

                #region "C# Patch Fix"
                // Patches solved in C# (below)
                /* // PowerShell Fix
                execute_program_show_cmd(@"C:\Windows\System32\cmd.exe", "/c netsh advfirewall import " + Path.GetTempPath() + extraction_folder_name + "\\WindowsUpdatesInstaller.Resources.DisableSMBv1.wfw");
                execute_program_show_cmd(@"C:\Windows\System32\cmd.exe", "/c netsh advfirewall set allprofiles state on");

                using (PowerShell ps = PowerShell.Create())
                {
                    Pipeline pipe = ps.Runspace.CreatePipeline();
                    pipe.Commands.AddScript("Set-SmbServerConfiguration -EnableSMB1Protocol $false");

                    try
                    {
                        Collection<PSObject> results = pipe.Invoke();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }*/
                #endregion
                #region "Embedded File Patch Fix"
                foreach (FileInfo f in d.GetFiles())
                {
                    if (f.Extension == ".msu")
                    {
                        // There shouldn't be any .msu files in this patch installer
                        /*// Enable Background Intelligent Transfer
                        execute_program_show_cmd("sc", "config BITS start= demand");
                        execute_program_show_cmd("net", "start BITS /y");

                        // Enable Windows Update
                        execute_program_show_cmd("sc", "config wuauserv start= demand");
                        execute_program_show_cmd("net", "start wuauserv /y");

                        // Install the .msu
                        execute_program_show_cmd("C:\\Windows\\System32\\wusa.exe", "\"" + f.FullName + "\"" + " /quiet /norestart");*/
                    }
                    else if (f.Extension == ".exe")
                    {

                    }
                    else if (f.Extension == ".bat")
                    {

                    }
                }
                #endregion
                #region "Delete the extracted files"
                foreach (FileInfo f in d.GetFiles()) f.Delete();
                #endregion
                #region "Update the progressBar and richTextBox"
                form.Invoke((MethodInvoker)delegate {
                    form.progressBar1.Value = form.progressBar1.Maximum;
                });
                #endregion

                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Environment.Exit(-1);
            }
            finally
            {
                Application.Exit();
            }
        }
        private static void uninstall()
        {
            try { 


                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Environment.Exit(-1);
            }
            finally
            {
                Application.Exit();
            }
        }

        #region "Extract Method"
        private static void ExtractEmbeddedResource(string outputDir, string resourceLocation, string file)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceLocation /*+ @"."*/ + file))
            {
                using (FileStream fileStream = new FileStream(System.IO.Path.Combine(outputDir, file), FileMode.Create))
                {
                    for (int i = 0; i < stream.Length; i++)
                    {
                        fileStream.WriteByte((byte)stream.ReadByte());
                    }
                    fileStream.Close();
                }
            }
        }
        #endregion
        #region "Add/Remove Directory Security"
        // Adds an ACL entry on the specified directory for the specified account.
        public static void AddDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            // Create a new DirectoryInfo object.
            DirectoryInfo dInfo = new DirectoryInfo(FileName);

            // Get a DirectorySecurity object that represents the 
            // current security settings.
            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            // Add the FileSystemAccessRule to the security settings. 
            dSecurity.AddAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            ControlType));

            // Set the new access settings.
            dInfo.SetAccessControl(dSecurity);

        }

        // Removes an ACL entry on the specified directory for the specified account.
        public static void RemoveDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            // Create a new DirectoryInfo object.
            DirectoryInfo dInfo = new DirectoryInfo(FileName);

            // Get a DirectorySecurity object that represents the 
            // current security settings.
            DirectorySecurity dSecurity = dInfo.GetAccessControl();

            // Add the FileSystemAccessRule to the security settings. 
            dSecurity.RemoveAccessRule(new FileSystemAccessRule(Account,
                                                            Rights,
                                                            ControlType));

            // Set the new access settings.
            dInfo.SetAccessControl(dSecurity);

        }
        #endregion
        #region "Run CMD"
        // Run a single command using the CMD. (CMD window is shown).
        public static int execute_program_show_cmd(string filename, string arguments)
        {
            // i.e. execute_program_show_cmd("C:\\Windows\\Sysnative\\manage-bde.exe", "-resume C:");
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = filename;
            startInfo.Arguments = arguments;
            process.StartInfo = startInfo;
            process.StartInfo.Verb = "runas";
            process.Start();
            process.WaitForExit();
            return process.ExitCode;
        }

        // Run a single command using the CMD. (CMD window is not shown).
        public static int execute_program_hide_cmd(string filename, string arguments)
        {
            // i.e. execute_program_hide_cmd("C:\\Windows\\Sysnative\\manage-bde.exe", "-resume C:");
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = filename;
            startInfo.Arguments = arguments;
            process.StartInfo = startInfo;
            process.StartInfo.Verb = "runas";
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = false;
            process.Start();
            process.WaitForExit();
            return process.ExitCode;
        }

        #endregion
        #region "OnProcessExit"
        static void OnProcessExit(object sender, EventArgs e)
        {
            try
            {
                // Delete the %temp%\3zx.xxxxx\ directories
                string[] files = Directory.GetDirectories(Path.GetTempPath(), "*.*");
                foreach (string f in files)
                {
                    if (f.StartsWith(Path.GetTempPath() + "3z"))
                    {
                        // Delete the %temp%\[random]\ directory
                        RemoveDirectorySecurity(f, System.Security.Principal.WindowsIdentity.GetCurrent().Name, FileSystemRights.Read, AccessControlType.Deny);
                        RemoveDirectorySecurity(f, System.Security.Principal.WindowsIdentity.GetCurrent().Name, FileSystemRights.ExecuteFile, AccessControlType.Allow);

                        string[] filess = Directory.GetFiles(f, "*.*");
                        foreach (string ff in filess)
                        {
                            File.Delete(ff);
                        }
                        Directory.Delete(f);
                    }
                }
            }
            catch (Exception ex) { }
        }
        #endregion
    }
}
