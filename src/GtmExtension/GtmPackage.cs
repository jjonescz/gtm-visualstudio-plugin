using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.ComponentModel;
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
        private string gtmExe, prevPath, status;
        private IVsStatusbar statusBar;
        private IVsEditorAdaptersFactoryService editor;
        private WindowEvents windowEvents;
        private DocumentEvents documentEvents;
        private IVsTextManager textManager;
        private IWpfTextView wpfTextView;
        private DateTime lastUpdate;
        private static readonly TimeSpan updateInterval = TimeSpan.FromSeconds(30.0);

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

        private async Task ShowErrorAsync(string message)
        {
            // Switch to UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            // Show message box.
            var uiShell = (IVsUIShell)await GetServiceAsync(typeof(SVsUIShell));
            if (uiShell == null) { return; }
            Guid clsid = Guid.Empty;
            ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(0, ref clsid, "GtmExtension", message, string.Empty, 0,
                OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST, OLEMSGICON.OLEMSGICON_CRITICAL, 0, out var result));
        }

        #endregion

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
            // Try to find executable `gtm`.
            if (ExistsOnPath("gtm"))
            {
                gtmExe = "gtm";
            }

            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Show error if we don't find `gtm` on PATH.
            if (gtmExe == null)
            {
                await ShowErrorAsync("We couldn't find gtm executable.");
                return;
            }

            // Verify version.
            if (!ExecuteForOutput(gtmExe, "verify \">= 1.1.0\"").Contains("true"))
            {
                await ShowErrorAsync("Old version of gtm is installed. Please install at least version 1.1.0");
                return;
            }

            // Get the status bar.
            statusBar = (IVsStatusbar)await GetServiceAsync(typeof(SVsStatusbar));
            if (statusBar == null) { throw new InvalidOperationException("No status bar."); }

            // Unfroze it if it's frozen.
            ErrorHandler.ThrowOnFailure(statusBar.IsFrozen(out var frozen));
            if (frozen != 0)
            {
                ErrorHandler.ThrowOnFailure(statusBar.FreezeOutput(0));
            }

            // Get DTE.
            var dte = (DTE)await GetServiceAsync(typeof(DTE));
            if (dte == null) { throw new InvalidOperationException("No DTE."); }

            // Get text manager.
            textManager = (IVsTextManager)await GetServiceAsync(typeof(SVsTextManager));
            if (textManager == null) { throw new InvalidOperationException("No TextManager."); }

            // Get editor adapters.
            var componentModel = (IComponentModel)await GetServiceAsync(typeof(SComponentModel));
            if (componentModel == null) { throw new InvalidOperationException("No ComponentModel."); }
            editor = componentModel.GetService<IVsEditorAdaptersFactoryService>();

            // Subscribe to events. We keep the events objects so that they don't get GC'ed.
            windowEvents = dte.Events.WindowEvents;
            windowEvents.WindowActivated += WindowEvents_WindowActivated;

            documentEvents = dte.Events.DocumentEvents;
            documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;

            Subscribe();
        }

        private void WindowEvents_WindowActivated(Window GotFocus, Window LostFocus)
        {
            Subscribe();
        }
        private string GetFilePath(ITextView textView)
        {
            return textView.TextBuffer.Properties.GetProperty<ITextDocument>(typeof(ITextDocument)).FilePath;
        }
        private void Subscribe()
        {
            // Unsubscribe the previously focused window.
            if (wpfTextView != null)
            {
                wpfTextView.Caret.PositionChanged -= Caret_PositionChanged;
                wpfTextView.LayoutChanged -= WpfTextView_LayoutChanged;
                wpfTextView = null;
            }

            // Subsribe the currently focused window.
            ErrorHandler.ThrowOnFailure(textManager.GetActiveView(0, null, out IVsTextView textView));
            wpfTextView = editor.GetWpfTextView(textView);
            wpfTextView.Caret.PositionChanged += Caret_PositionChanged;
            wpfTextView.LayoutChanged += WpfTextView_LayoutChanged;

            Update(GetFilePath(wpfTextView));
        }

        private void WpfTextView_LayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            Update(GetFilePath(wpfTextView));
        }
        private void Caret_PositionChanged(object sender, CaretPositionChangedEventArgs e)
        {
            Update(GetFilePath(e.TextView));
        }
        private void DocumentEvents_DocumentSaved(Document Document)
        {
            Update(Document.FullName);
        }
        private void Update(string path)
        {
            var time = DateTime.Now;
            if (time - lastUpdate >= updateInterval ||
                path != prevPath)
            {
                status = ExecuteForOutput(gtmExe, $"record --status \"{path}\"");
                if (!string.IsNullOrWhiteSpace(status))
                {
                    statusBar.SetText($"GTM: {status}*");
                }

                prevPath = path;
            }
            else if (!string.IsNullOrWhiteSpace(status))
            {
                statusBar.SetText($"GTM: {status}");
            }
            lastUpdate = time;
        }
        #endregion
    }
}
