using DocFetcherSelection.Properties;
using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace DocFetchSelector
{
    public class Addon : CitaviAddOn<MainForm>
    {
        #region Constants

        const string DocFetcherSelectionAddon_Command = "DocFetcherSelectionAddon.Command";

        #endregion

        #region Methods

        public override void OnHostingFormLoaded(MainForm mainForm)
        {
            if (mainForm.Project.ProjectType == ProjectType.DesktopSQLite)
            {
                mainForm.GetMainCommandbarManager()
                    .GetReferenceEditorCommandbar(MainFormReferenceEditorCommandbarId.Menu)
                    .GetCommandbarMenu(MainFormReferenceEditorCommandbarMenuId.Edit)
                    .AddCommandbarButton(DocFetcherSelectionAddon_Command, Strings.Command);
            }

            base.OnHostingFormLoaded(mainForm);
        }

        public override void OnBeforePerformingCommand(MainForm mainForm, BeforePerformingCommandEventArgs e)
        {
            if (e.Key.Equals(DocFetcherSelectionAddon_Command, StringComparison.OrdinalIgnoreCase))
            {
                e.Handled = true;
                var data = Clipboard.GetDataObject().GetData(typeof(string)) as string;
                if (!string.IsNullOrEmpty(data))
                {
                    Run(mainForm,data.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries));
                }
            }

            base.OnBeforePerformingCommand(mainForm, e);
        }

        public override void OnLocalizing(MainForm form)
        {
            var button = form.GetMainCommandbarManager()
                             .GetReferenceEditorCommandbar(MainFormReferenceEditorCommandbarId.Menu)
                             .GetCommandbarMenu(MainFormReferenceEditorCommandbarMenuId.Edit)
                             .GetCommandbarButton(DocFetcherSelectionAddon_Command);

            if (button != null) button.Text = Strings.Command;

            base.OnLocalizing(form);
        }

        void Run(MainForm mainForm, IEnumerable<string> clipboardFiles)
        {
          
            var foundedReferences = new List<Reference>();
            try
            {
                foreach (var reference in mainForm.Project.References)
                {
                    foreach (var location in reference.Locations.ToList())
                    {
                        var path = location.Address.Resolve().GetLocalPathSafe();
                        if (clipboardFiles.Contains(path))
                        {
                            foundedReferences.Add(reference);
                            break;
                        }
                    }
                }

                if (foundedReferences.Any())
                {
                    var filter = new ReferenceFilter(foundedReferences, "References with locations from file", false);
                    mainForm.ReferenceEditorFilterSet.Filters.ReplaceBy(new List<ReferenceFilter> { filter });
                }
            }

            finally
            {
                if (!foundedReferences.Any())
                {
                    MessageBox.Show(Strings.FoundNoMatches, Program.ActiveProjectShell.PrimaryMainForm.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        #endregion
    }
}
