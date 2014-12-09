﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace manypdftoone
{
    public class ManyPdfToOne
    {
        private DirectoryInfo _sourceDir = null;
        private DirectoryInfo _destinationDir = null;
        private FileInfo[] _sourceFiles = null;
        private FileInfo _destinationFile = null;
        private eMergePagesMode _mergeMode = eMergePagesMode.OneSide;
        private List<SourceFile> _sourceFilesDef = new List<SourceFile>();
        private string _gsfileName = string.Empty;

        /// <summary>
        /// Initialize class with instantiate all objects.
        /// </summary>
        /// <param name="sourceDir">Source directory for find all pdf files</param>
        /// <param name="destinationDir"></param>
        public ManyPdfToOne(string sourceDir, string destinationDir, eMergePagesMode mergeMode)
        {
            _sourceDir = new DirectoryInfo(sourceDir);
            _destinationDir = new DirectoryInfo(destinationDir);

            _sourceFiles = _sourceDir.GetFiles("*.pdf", SearchOption.AllDirectories);
            _destinationFile = new FileInfo(Path.GetTempFileName());
        }

        /// <summary>
        /// 
        /// </summary>
        public void Merge()
        {
            string gsFileName = string.Empty;

            RegistryKey gsKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\GPL Ghostscript");
            string[] gssKeys = gsKey.GetSubKeyNames();
            if (gssKeys.Length == 0)
                throw new Exception("Ghostscript nie jest zainstalowany.");

            gsKey = gsKey.OpenSubKey(gssKeys[0]);
            FileInfo gsLibDir = new FileInfo(gsKey.GetValue("GS_DLL").ToString());

            gsFileName = gsLibDir.DirectoryName;

            if (Common.is64BitProcess)
                gsFileName = Path.Combine(gsFileName, "GSWIN64C.EXE");
            else
                gsFileName = Path.Combine(gsFileName, "GSWIN32C.EXE");

            if (!new FileInfo(gsFileName).Exists)
                throw new FileNotFoundException(string.Format("Brak pliku {0}", gsFileName));

            _gsfileName = gsFileName;

            for (int i = 0; i < _sourceFiles.Length; i++)
            {
                SourceFile sf = new SourceFile();
                sf.FilePath = _sourceFiles[i].FullName;
                sf.PagesCount = CountPages(sf.FilePath);
                _sourceFilesDef.Add(sf);
                if (sf.PagesCount % 2 != 0)
                    _sourceFilesDef.Add(GetOddPage());
            }
        }

        private SourceFile GetOddPage()
        {
            SourceFile sf = new SourceFile();
            sf.FilePath = Path.GetTempFileName();
            ExtractSaveResource("manypdftoone.Resources.blank.pdf", sf.FilePath);
            sf.PagesCount = CountPages(sf.FilePath);
            return sf;
        }

        private int CountPages(string filePath)
        {
            int pc = -1;
            ShellProcess sp = new ShellProcess();
            string[] response = sp.Run(_gsfileName
                , string.Format("-q -c \"({0}) (r) file runpdfbegin pdfpagecount = quit\" -dNODISPLAY", filePath.Replace("\\", "/")));

            if (string.IsNullOrEmpty(response[1]))
            {
                if (!int.TryParse(response[0], out pc))
                    throw new Exception(string.Format("nie pobrano ilości stron dla {0}", filePath));
            }
            return pc;
        }

        private void ExtractSaveResource(String filename, String location)
        {
            System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
            Stream resFilestream = a.GetManifestResourceStream(filename);
            if (resFilestream != null)
            {
                BinaryReader br = new BinaryReader(resFilestream);
                FileStream fs = new FileStream(location, FileMode.Create); // say 
                BinaryWriter bw = new BinaryWriter(fs);
                byte[] ba = new byte[resFilestream.Length];
                resFilestream.Read(ba, 0, ba.Length);
                bw.Write(ba);
                br.Close();
                bw.Close();
                resFilestream.Close();
            }
        }
    }
}
