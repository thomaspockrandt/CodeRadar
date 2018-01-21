using EnvDTE;
using System;
using System.Linq;
using System.Media;
using System.Speech.Synthesis;
using System.Text.RegularExpressions;

namespace CodeRadar
{
    public class CommandBase
    {
        public void PlaySound()
        {
            var mediaFile = Environment.ExpandEnvironmentVariables(@"%windir%\Media\chimes.wav");
            using (var soundPlayer = new SoundPlayer(mediaFile))
            {
                soundPlayer.Play();
            }
        }

        public string GetCurrentLineText(DTE dte)
        {
            var selection = (TextSelection)dte.ActiveDocument.Selection;

            var activePoint = selection.ActivePoint;
            selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
            selection.EndOfLine(true);
            var line = selection.Text;

            selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
            selection.Collapse();

            return line;
        }

        internal void InvokeEditPreviousMethod(DTE dte)
        {            
            dte.ExecuteCommand("Edit.PreviousMethod");
        }

        internal void InvokeEditNextMethod(DTE dte)
        {
            dte.ExecuteCommand("Edit.NextMethod");
        }

        internal string GetMethodName(string line)
        {
            var group = Regex.Match(line, @"\s([.\S]*)\(").Groups[1];

            return string.IsNullOrWhiteSpace(group.Value) ? line : group.Value;
        }

        internal void SpeakText(string text)
        {
            var synth = new SpeechSynthesizer();

            synth.SetOutputToDefaultAudioDevice();

            synth.SpeakAsync(text);
        }
    }
}
