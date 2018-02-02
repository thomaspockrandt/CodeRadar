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
        private DTE dte;

        public CommandBase(DTE dte)
        {
            this.dte = dte;
        }

        public void PlaySound()
        {
            var mediaFile = Environment.ExpandEnvironmentVariables(@"%windir%\Media\chimes.wav");
            using (var soundPlayer = new SoundPlayer(mediaFile))
            {
                soundPlayer.Play();
            }
        }

        public string GetCurrentLineText()
        {
            var selection = (TextSelection)this.dte.ActiveDocument.Selection;

            var activePoint = selection.ActivePoint;
            selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
            selection.EndOfLine(true);
            var line = selection.Text;

            selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
            selection.Collapse();

            return line;
        }

        internal void InvokeEditPreviousMethod()
        {            
            this.dte.ExecuteCommand("Edit.PreviousMethod");
        }

        internal void InvokeEditNextMethod()
        {
            this.dte.ExecuteCommand("Edit.NextMethod");
        }

        internal string GetMethodName(string line)
        {
            var group = Regex.Match(line, @"\s([.\S]*)\(").Groups[1];

            var name = string.IsNullOrWhiteSpace(group.Value) ? line : group.Value;

            return name;
        }

        internal string GetCleanLineOfCode(string line)
        {
            var group = Regex.Match(line, @"\s([.\S]*)\(").Groups[1];

            var cleanLine = string.IsNullOrWhiteSpace(group.Value) ? line : group.Value;

            cleanLine = cleanLine.Replace(".", " ");

            return cleanLine;
        }

        internal void SpeakText(string text)
        {
            var synth = new SpeechSynthesizer();

            synth.SetOutputToDefaultAudioDevice();

            synth.SpeakAsync(text);
        }
    }
}
