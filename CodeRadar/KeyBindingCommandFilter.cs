using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text.Editor;
using System.Diagnostics;

namespace CodeRadar
{
    internal class KeyBindingCommandFilter : IOleCommandTarget
    {
        private IWpfTextView m_textView;
        internal IOleCommandTarget m_nextTarget;
        internal bool m_added;
        internal bool m_adorned;

        public KeyBindingCommandFilter(IWpfTextView textView)
        {
            m_textView = textView;
            m_adorned = false;
        }

        int IOleCommandTarget.QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return m_nextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        int IOleCommandTarget.Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            //if (m_adorned == false)
            //{
            //    char typedChar = char.MinValue;
            //    if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
            //    {
            //        typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
            //        if (typedChar.Equals('+'))
            //        {
            //            new PurpleCornerBox(m_textView);
            //            m_adorned = true;
            //        }
            //    }
            //}
            return m_nextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }


    }
}
