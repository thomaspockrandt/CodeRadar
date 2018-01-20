using EnvDTE;
using Microsoft.VisualStudio.OLE.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace CodeRadar
{
    public static class DteHelpers
    {
        /// <summary>
        /// Returns an IVsTextView for the given file path, if the given file is open in Visual Studio.
        /// </summary>
        /// <param name="filePath">Full Path of the file you are looking for.</param>
        /// <returns>The IVsTextView for this file, if it is open, null otherwise.</returns>
        internal static Microsoft.VisualStudio.TextManager.Interop.IVsTextView GetIVsTextView(DTE dte, string filePath)
        {
            Microsoft.VisualStudio.OLE.Interop.IServiceProvider sp = (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)dte;
            Microsoft.VisualStudio.Shell.ServiceProvider serviceProvider = new Microsoft.VisualStudio.Shell.ServiceProvider(sp);

            Microsoft.VisualStudio.Shell.Interop.IVsUIHierarchy uiHierarchy;
            uint itemID;
            Microsoft.VisualStudio.Shell.Interop.IVsWindowFrame windowFrame;
            Microsoft.VisualStudio.Text.Editor.IWpfTextView wpfTextView = null;
            if (Microsoft.VisualStudio.Shell.VsShellUtilities.IsDocumentOpen(serviceProvider, filePath, Guid.Empty, out uiHierarchy, out itemID, out windowFrame))
            {
                // Get the IVsTextView from the windowFrame.
                return Microsoft.VisualStudio.Shell.VsShellUtilities.GetTextView(windowFrame);
            }

            return null;
        }

        public static DTE GetActiveInstance()
        {
            var instance = Marshal.GetActiveObject("VisualStudio.DTE.15.0") as DTE;

            return instance;
        }

        public static IEnumerable<DTE> GetInstances()
        {
            IRunningObjectTable rot;
            IEnumMoniker enumMoniker;
            int retVal = GetRunningObjectTable(0, out rot);

            if (retVal == 0)
            {
                rot.EnumRunning(out enumMoniker);

                IMoniker[] moniker = new IMoniker[1];
                uint fetched;
                while (enumMoniker.Next(1, moniker, out fetched) == 0)
                {
                    IBindCtx bindCtx;
                    CreateBindCtx(0, out bindCtx);
                    string displayName;
                    moniker[0].GetDisplayName(bindCtx, null, out displayName);
                    Console.WriteLine("Display Name: {0}", displayName);
                    bool isVisualStudio = displayName.StartsWith("!VisualStudio");
                    if (isVisualStudio)
                    {
                        object ppunkObject;
                        IMoniker m = moniker[0];
                        rot.GetObject(m, out ppunkObject);
                        yield return (DTE)ppunkObject;
                    }
                }
            }
        }

        [DllImport("ole32.dll")]
        private static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
}
