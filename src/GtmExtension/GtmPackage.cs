using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;
using Process = System.Diagnostics.Process;
using Task = System.Threading.Tasks.Task;

namespace GtmExtension
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
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)] // Load the extension when a solution is open.
    public sealed class GtmPackage : AsyncPackage
    {
        /// <summary>
        /// GtmPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "208330cb-08c8-4105-b9b2-8f010fbeaf49";

        /// <summary>
        /// Initializes a new instance of the <see cref="GtmPackage"/> class.
        /// </summary>
        public GtmPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
        }

        #endregion
    }

    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    public sealed class AnotherListener : IVsTextViewCreationListener
    {
        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
        }
    }

    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    public sealed class GtmListener : IVsTextViewCreationListener
    {
        private bool initialized;
        private string gtmExe, prevPath, status;
        private DateTime lastUpdate;
        private IVsEditorAdaptersFactoryService editor;
        private DocumentEvents documentEvents;
        private DTE dte;
        private IComponentModel componentModel;
        private IVsStatusbar statusbar;
        private IVsUIShell uiShell;
        private static readonly TimeSpan updateInterval = TimeSpan.FromSeconds(30.0);

        #region Imports
        [Import]
        public SVsServiceProvider ServiceProvider { get; set; }
        #endregion

        #region Helper Functions

        private static Process ExecuteProcess(string exeName, string arguments)
        {
            try
            {
                var p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.FileName = exeName;
                p.StartInfo.Arguments = arguments;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.Start();
                p.WaitForExit();
                return p;
            }
            catch (Win32Exception)
            {
                throw new Exception($"Executable \"{exeName}\" was not found.");
            }
        }
        private static int Execute(string exeName, string arguments = null)
        {
            return ExecuteProcess(exeName, arguments).ExitCode;
        }
        private static string ExecuteForOutput(string exeName, string arguments = null)
        {
            Process p = ExecuteProcess(exeName, arguments);
            return p.StandardOutput.ReadToEnd();
        }

        /// <summary>
        /// Checks if executable <paramref name="exeName"/> can be found in system's PATH.
        /// </summary>
        /// <seealso href="https://stackoverflow.com/a/24405838/9080566"/>
        private static bool ExistsOnPath(string exeName)
        {
            return Execute("where", exeName) == 0;
        }

        private void ShowError(string message)
        {
            // Show message box.
            Guid clsid = Guid.Empty;
            ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(0, ref clsid, "GtmExtension", message, string.Empty, 0,
                OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST, OLEMSGICON.OLEMSGICON_CRITICAL, 0, out var result));
        }

        private string GetFilePath(ITextView textView)
        {
            return textView.TextBuffer.Properties.GetProperty<ITextDocument>(typeof(ITextDocument)).FilePath;
        }

        private void GetService<T>(ref T field) { field = GetService<T>(); }
        private T GetService<T>() => GetService<T, T>();
        private I GetService<S, I>()
        {
            var service = ServiceProvider.GetService(typeof(S));
            if (service == null) { throw new InvalidOperationException($"No {typeof(S).Name}."); }
            return (I)service;
        }

        #endregion

        #region Event handlers

        public void VsTextViewCreated(IVsTextView textView)
        {
            Initialize();

            var wpfTextView = editor.GetWpfTextView(textView);
            wpfTextView.LayoutChanged += WpfTextView_LayoutChanged;
            wpfTextView.Caret.PositionChanged += Caret_PositionChanged;

            Update(GetFilePath(wpfTextView));
        }
        private void WpfTextView_LayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            Update(GetFilePath((ITextView)sender));
        }
        private void Caret_PositionChanged(object sender, CaretPositionChangedEventArgs e)
        {
            Update(GetFilePath(e.TextView));
        }
        private void DocumentEvents_DocumentSaved(Document Document)
        {
            Update(Document.FullName, force: true);
        }

        #endregion

        private void Initialize()
        {
            if (initialized) { return; }
            initialized = true;

            // Get services.
            GetService(ref uiShell);
            GetService(ref statusbar);
            componentModel = GetService<SComponentModel, IComponentModel>();
            GetService(ref dte);
            editor = componentModel.GetService<IVsEditorAdaptersFactoryService>();

            // Try to find executable `gtm`.
            if (ExistsOnPath("gtm"))
            {
                gtmExe = "gtm";
            }
            else
            {
                ShowError("We couldn't find gtm executable.");
                return;
            }

            // Verify version.
            if (!ExecuteForOutput(gtmExe, "verify \">= 1.1.0\"").Contains("true"))
            {
                ShowError("Old version of gtm is installed. Please install at least version 1.1.0");
                return;
            }

            // Unfroze status bar if it's frozen.
            ErrorHandler.ThrowOnFailure(statusbar.IsFrozen(out var frozen));
            if (frozen != 0)
            {
                ErrorHandler.ThrowOnFailure(statusbar.FreezeOutput(0));
            }

            // Subscribe to events. We keep the events object so that it doesn't get GC'ed.
            documentEvents = dte.Events.DocumentEvents;
            documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;
        }
        private void Update(string path, bool force = false)
        {
            var time = DateTime.Now;
            if (force ||
                time - lastUpdate >= updateInterval ||
                path != prevPath)
            {
                status = ExecuteForOutput(gtmExe, $"record --status \"{path}\"");
                if (!string.IsNullOrWhiteSpace(status))
                {
                    statusbar.SetText($"GTM: {status}*");
                }

                prevPath = path;
            }
            else if (!string.IsNullOrWhiteSpace(status))
            {
                statusbar.SetText($"GTM: {status}");
            }
            lastUpdate = time;
        }
    }
}
