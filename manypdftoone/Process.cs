using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.IO;

namespace manypdftoone
{
    public class ShellProcess
    {
        public string[] Run(string file, string args)
        {
            string ret = string.Empty;
            string err = string.Empty;
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = args;
            psi.CreateNoWindow = true;
            psi.FileName = file;
            psi.LoadUserProfile = false;
            psi.RedirectStandardError = true;
            psi.RedirectStandardInput = false;
            psi.RedirectStandardOutput = true;
            psi.UseShellExecute = false;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.WorkingDirectory = new FileInfo(file).DirectoryName;

            Process p = new Process();
            p.StartInfo = psi;
            p.Start();
            p.WaitForExit();
            ret = p.StandardOutput.ReadToEnd();
            err = p.StandardError.ReadToEnd();
            p.Dispose();

            return new string[] { ret, err };
        }
    }
}
