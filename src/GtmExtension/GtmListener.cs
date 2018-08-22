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
using Process = System.Diagnostics.Process;

namespace GtmExtension
{
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
