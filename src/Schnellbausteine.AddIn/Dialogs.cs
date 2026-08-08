using System;
using System.Drawing;
using System.Windows.Forms;
using Schnellbausteine.Core;

namespace Schnellbausteine.AddIn
{
    internal sealed class WindowHandleOwner : IWin32Window
    {
        public WindowHandleOwner(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; private set; }
    }

    internal sealed class NameDialog : Form
    {
        private readonly TextBox nameBox;

        public NameDialog(string title, string initialName)
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(390, 115);

            var label = new Label { Left = 12, Top = 14, Width = 360, Text = "Name:" };
            nameBox = new TextBox { Left = 12, Top = 36, Width = 365, Text = initialName ?? string.Empty, MaxLength = 100 };
            var ok = new Button { Left = 221, Top = 75, Width = 75, Text = "OK", DialogResult = DialogResult.OK };
            var cancel = new Button { Left = 302, Top = 75, Width = 75, Text = "Abbrechen", DialogResult = DialogResult.Cancel };
            AcceptButton = ok;
            CancelButton = cancel;
            Controls.AddRange(new Control[] { label, nameBox, ok, cancel });
        }

        public string SnippetName { get { return nameBox.Text.Trim(); } }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            nameBox.SelectAll();
            nameBox.Focus();
        }
    }

    internal sealed class ManageDialog : Form
    {
        private readonly SnippetRepository repository;
        private readonly Action changed;
        private readonly ListBox list;

        public ManageDialog(SnippetRepository repository, Action changed)
        {
            this.repository = repository;
            this.changed = changed;
            Text = "Schnellbausteine verwalten";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(500, 330);
            Size = new Size(600, 420);

            list = new ListBox { Dock = DockStyle.Fill, DisplayMember = "Name" };
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 46, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(6) };
            var close = new Button { Text = "Schließen", Width = 90, DialogResult = DialogResult.OK };
            var delete = new Button { Text = "Löschen", Width = 90 };
            var rename = new Button { Text = "Umbenennen", Width = 100 };
            delete.Click += Delete;
            rename.Click += Rename;
            buttons.Controls.AddRange(new Control[] { close, delete, rename });
            Controls.Add(list);
            Controls.Add(buttons);
            AcceptButton = close;
            LoadItems();
        }

        private Snippet Selected { get { return list.SelectedItem as Snippet; } }

        private void LoadItems()
        {
            string selectedId = Selected == null ? null : Selected.Id;
            list.DataSource = null;
            list.DataSource = repository.GetAll();
            if (selectedId != null)
            {
                foreach (object item in list.Items)
                {
                    if (((Snippet)item).Id == selectedId) list.SelectedItem = item;
                }
            }
        }

        private void Rename(object sender, EventArgs e)
        {
            Snippet item = Selected;
            if (item == null) return;
            using (var dialog = new NameDialog("Schnellbaustein umbenennen", item.Name))
            {
                while (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        repository.Rename(item.Id, dialog.SnippetName);
                        LoadItems();
                        changed();
                        return;
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show(this, exception.Message, "Umbenennen nicht möglich", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void Delete(object sender, EventArgs e)
        {
            Snippet item = Selected;
            if (item == null) return;
            if (MessageBox.Show(this, "Schnellbaustein \"" + item.Name + "\" wirklich löschen?", "Löschen bestätigen",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            repository.Delete(item.Id);
            LoadItems();
            changed();
        }
    }
}
