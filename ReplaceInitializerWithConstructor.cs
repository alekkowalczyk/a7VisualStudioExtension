using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Text;
using EnvDTE;

namespace a7VisualStudioExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ReplaceInitializerWithConstructor
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("706dd8c4-1dce-45ca-b62f-271fb663a16a");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ReplaceInitializerWithConstructor"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ReplaceInitializerWithConstructor(Package package)
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
        public static ReplaceInitializerWithConstructor Instance
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
            Instance = new ReplaceInitializerWithConstructor(package);
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
            replaceInitializerWithConstructor();
            //string message = getSelectedText_V1();// string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            //string title = "Command1";

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        //private string getSelectedText_V1()
        //{
        //    string selectedText = string.Empty;
        //    DTE dte = this.ServiceProvider.GetService(typeof(DTE)) as DTE;

        //    Document doc = dte.ActiveDocument;
        //    var selection = doc.Selection as TextSelection;
        //    selection.ReplaceText(selection.Text, "blabla");
        //    return selection.Text;
        //}

        private void replaceInitializerWithConstructor()
        {
            DTE dte = this.ServiceProvider.GetService(typeof(DTE)) as DTE;

            Document doc = dte.ActiveDocument;
            var selection = doc.Selection as TextSelection;
            var text = selection.Text;
            if (string.IsNullOrEmpty(text))
            {
                showMessageBox("No selection");
                return;
            }
            if (!text.Trim().StartsWith("{"))
            {
                showMessageBox("Selection needs to start with {");
                return;
            }
            if (!text.Contains("}"))
            {
                showMessageBox("Selection needs to contain } char");
                return;
            }
            var sb = new StringBuilder();
            var isInsideInitializer = false;
            var isInsidePropertyName = false;
            var isInsidePropertySetter = false;
            for (var i = 0; i < text.Length; i++)
            {
                var ch = text[i];
                var nextCh = i < text.Length - 1 ? text[i + 1] : char.MinValue;
                if (!isInsideInitializer)
                {
                    if (ch == '{')
                    {
                        isInsideInitializer = true;
                        sb.Append('(');
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }
                else
                {
                    if (ch == '}')
                    {
                        sb.Append(')');
                        isInsideInitializer = false;
                    }
                    else if (isInsideInitializer && !isInsidePropertyName && !isInsidePropertySetter && char.IsLetterOrDigit(ch))
                    {
                        isInsidePropertyName = true;
                        sb.Append(char.ToLower(ch));
                    }
                    else if (isInsidePropertyName && nextCh == '=')
                    {
                        sb.Append(':');
                        sb.Append(' ');
                        i++;
                        isInsidePropertyName = false;
                        isInsidePropertySetter = true;
                    }
                    else if (isInsidePropertySetter && ch == ',')
                    {
                        isInsidePropertySetter = false;
                        sb.Append(ch);
                    }
                    else
                    {
                        sb.Append(ch);
                    }
                }
            }
            selection.ReplaceText(selection.Text, sb.ToString());
        }

        private void showMessageBox(string message)
        {
            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                "a7VSExtension",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
