using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using EnvDTE;
using EnvDTE80;
using System.IO;
using System.Text;
using Task = System.Threading.Tasks.Task;

namespace ForceAllUtf8
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.1", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(AllForceUtf8Package.PackageGuidString)]
    //[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class AllForceUtf8Package : AsyncPackage
    {
        /// <summary>
        /// AllForceUtf8Package GUID string.
        /// </summary>
        public const string PackageGuidString = "e3f46549-8a3c-44f5-b00d-6219c563479a";

        /// <summary>
        /// Initializes a new instance of the <see cref="AllForceUtf8Package"/> class.
        /// </summary>
        public AllForceUtf8Package()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }

        #region Package Members

        private DocumentEvents documentEvents;
        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            //base.Initialize();

            var dte = GetService(typeof(DTE)) as DTE2;
            documentEvents = dte.Events.DocumentEvents;
            documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;
        }
        void DocumentEvents_DocumentSaved(Document document)
        {
            if (document.Kind != "{8E7B96A8-E33D-11D0-A6D5-00C04FB67F6A}")
            {
                // then it's not a text file
                return;
            }

            string path = document.FullName;
            bool isJava = false;
            if (path.EndsWith(".java"))
            {
                Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "file '{0}' is java file, will convert to utf-8(no BOM).", path));
                isJava = true;
            }

            try
            {
                var stream = new FileStream(path, FileMode.Open);
                var reader = new StreamReader(stream, Encoding.Default, true);
                reader.Read();

                var preambleBytes = reader.CurrentEncoding.GetPreamble();
                if (preambleBytes.Length == 3 && 
                    preambleBytes[0] == 0xEF && preambleBytes[1] == 0xBB && preambleBytes[2] == 0xBF &&
                    reader.CurrentEncoding.EncodingName == Encoding.UTF8.EncodingName)
                {
                    stream.Close();
                    return;
                }

                string text;

                try
                {
                    stream.Position = 0;
                    reader = new StreamReader(stream, new UTF8Encoding(true, true));
                    text = reader.ReadToEnd();
                    stream.Close();
                    File.WriteAllText(path, text, new UTF8Encoding(!isJava));  // java file save with no-BOM
                    Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Already convert file '{0}' encoding to utf-8(BOM).", path));
                }
                catch (DecoderFallbackException)
                {
                    stream.Position = 0;
                    reader = new StreamReader(stream, Encoding.Default);
                    text = reader.ReadToEnd();
                    stream.Close();
                    File.WriteAllText(path, text, new UTF8Encoding(!isJava));
                    Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Already convert file '{0}' encoding to utf-8(BOM).", path));
                }

            }
            catch (Exception e)
            {
                Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Exception occured when converting file '{0}' encoding to utf-8(BOM):\n{1}", 
                    path, e.ToString()));
            }
        }

        #endregion
    }
}
