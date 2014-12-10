using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace manypdftoone
{
    public class ManyPdfToOne
    {
        private DirectoryInfo _sourceDir = null;
        private FileInfo[] _sourceFiles = null;
        private FileInfo _destinationFile = null;
        private eMergePagesMode _mergeMode = eMergePagesMode.OneSide;
        private List<SourceFile> _sourceFilesDef = new List<SourceFile>();
        private string _gsfileName = string.Empty;
        private Regex _rxFileName = new Regex("^[^\\/\\\\:\\*\\?\"\\<\\>\\|()]+$", RegexOptions.IgnoreCase);

        /// <summary>
        /// Initialize class with instantiate all objects.
        /// </summary>
        /// <param name="sourceDir">Source directory for find all pdf files</param>
        /// <param name="destinationFile"></param>
        public ManyPdfToOne(string sourceDir, string destinationFile, eMergePagesMode mergeMode)
        {
            Init(sourceDir, destinationFile, mergeMode);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceDir"></param>
        /// <param name="destinationDir"></param>
        /// <param name="mergeMode"></param>
        private void Init(string sourceDir, string destinationFile, eMergePagesMode mergeMode)
        {
            _sourceDir = new DirectoryInfo(sourceDir);
            _sourceFiles = _sourceDir.GetFiles("*.pdf", SearchOption.AllDirectories);
            File.Create(destinationFile).Close();
            _destinationFile = new FileInfo(destinationFile);
            _mergeMode = mergeMode;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Merge()
        {
            #region check ghostscript
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
            #endregion

            #region Prepare files definition
            var sortedList = from file in _sourceFiles.ToList()
                             orderby file.FullName ascending
                             select file;
            _sourceFiles = sortedList.ToArray<FileInfo>();

            for (int i = 0; i < _sourceFiles.Length; i++)
            {
                SourceFile sf = new SourceFile();
                sf.FilePath = _sourceFiles[i].FullName.Replace("\\", "/").Trim();
                sf.PagesCount = CountPages(sf.FilePath);

                if (!_rxFileName.IsMatch(new FileInfo(sf.FilePath).Name))
                    throw new Exception(string.Format("Nieprawidłowa nazwa pliku {0}. Niedozwolone znaki space^/\\:*?\"<>|()", new FileInfo(sf.FilePath).Name));

                _sourceFilesDef.Add(sf);

                if (_mergeMode == eMergePagesMode.TwoSide && sf.PagesCount % 2 != 0)
                    _sourceFilesDef.Add(GetEvenPage());
            } 
            #endregion

            #region building merge command
            StringBuilder gsMergeCommand = new StringBuilder()
                .Append(" -dBATCH")
                .Append(" -dNOPAUSE")
                .Append(" -dQUIET")
                .Append(" -dSHORTERRORS")
                .Append(" -sDEVICE=pdfwrite")
                .Append(" -sOutputFile=\"").Append(_destinationFile.FullName.Replace("\\", "/")).Append("\"");

            foreach (SourceFile sf in _sourceFilesDef)
                gsMergeCommand.Append(" \"").Append(sf.FilePath).Append("\""); 
            #endregion

            Console.WriteLine(gsMergeCommand.ToString());

            RunMerge(gsMergeCommand.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceDir"></param>
        /// <param name="destinationDir"></param>
        /// <param name="mergeMode"></param>
        public void Merge(string sourceDir, string destinationDir, eMergePagesMode mergeMode)
        {
            Init(sourceDir, destinationDir, mergeMode);
            Merge();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="command"></param>
        private void RunMerge(string command)
        {
            ShellProcess sp = new ShellProcess();
            string[] response = sp.Run(_gsfileName, command);

            Console.Write(response[0]);

            if (response[1].ToLower().Contains("error:"))
                throw new Exception(response[0] + response[1]);
        }

        private SourceFile GetEvenPage()
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
