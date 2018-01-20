using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Media;
using System.Speech.Synthesis;
using System.Text.RegularExpressions;
using System.Windows.Media;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace CodeRadar
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class MethodNavigationCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("4f19b86a-44d5-488e-8920-c8e4e51113b4");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="MethodNavigationCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private MethodNavigationCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static MethodNavigationCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new MethodNavigationCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = DteHelpers.GetInstances().First();
            dte.ExecuteCommand("Edit.NextMethod");

            //using (var soundPlayer = new SoundPlayer(@"c:\Windows\Media\chimes.wav"))
            //{
            //    soundPlayer.Play();
            //}

            var selection = (TextSelection)dte.ActiveDocument.Selection;
            var activePoint = selection.ActivePoint;
            selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
            selection.EndOfLine(true);
            var line = selection.Text;

            selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText);
            selection.Collapse();
            
            var group = Regex.Match(line, @"\s([.\S]*)\(").Groups[1];
            var synth = new SpeechSynthesizer();
            synth.SetOutputToDefaultAudioDevice();
            synth.SpeakAsync(group.Value);

            var fullPath = dte.ActiveDocument.FullName;
            var view = DteHelpers.GetIVsTextView(dte, fullPath);

            var brush = new SolidColorBrush(Colors.BlueViolet);
            brush.Freeze();
            var penBrush = new SolidColorBrush(Colors.Red);
            penBrush.Freeze();
            var pen = new Pen(penBrush, 0.5);
            pen.Freeze();

            // Draw a square with the created brush and pen
            var r = new System.Windows.Rect(0, 0, dte.ActiveWindow.Width, dte.ActiveWindow.Height);
            var geometry = new RectangleGeometry(r);
            var drawing = new GeometryDrawing(brush, pen, geometry);
            drawing.Freeze();

            var drawingImage = new DrawingImage(drawing);
            drawingImage.Freeze();

            view.im = new Image
            {
                Source = drawingImage,
            };

        }


    }
}
