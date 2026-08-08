using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
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

    internal enum SnippetMenuEntryKind
    {
        Snippet,
        Empty,
        Separator,
        Manage
    }

    internal sealed class SnippetMenuEntry
    {
        private SnippetMenuEntry(
            SnippetMenuEntryKind kind,
            string text,
            Snippet snippet,
            bool enabled)
        {
            Kind = kind;
            Text = text;
            Snippet = snippet;
            Enabled = enabled;
        }

        public SnippetMenuEntryKind Kind { get; private set; }
        public string Text { get; private set; }
        public Snippet Snippet { get; private set; }
        public bool Enabled { get; private set; }

        public static IList<SnippetMenuEntry> Create(IList<Snippet> snippets)
        {
            if (snippets == null) throw new ArgumentNullException("snippets");
            var entries = snippets
                .OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                .Select(item => new SnippetMenuEntry(
                    SnippetMenuEntryKind.Snippet,
                    item.Name,
                    item,
                    true))
                .ToList();
            if (entries.Count == 0)
            {
                entries.Add(new SnippetMenuEntry(
                    SnippetMenuEntryKind.Empty,
                    "Keine Schnellbausteine gespeichert",
                    null,
                    false));
            }
            entries.Add(new SnippetMenuEntry(
                SnippetMenuEntryKind.Separator,
                string.Empty,
                null,
                false));
            entries.Add(new SnippetMenuEntry(
                SnippetMenuEntryKind.Manage,
                "Schnellbausteine verwalten",
                null,
                true));
            return entries;
        }

        public static int FindEnabledIndex(
            IList<SnippetMenuEntry> entries,
            int selectedIndex,
            int lastEnabledIndex)
        {
            if (entries == null) throw new ArgumentNullException("entries");
            int direction = selectedIndex > lastEnabledIndex ? 1 : -1;
            for (int index = selectedIndex + direction;
                index >= 0 && index < entries.Count;
                index += direction)
            {
                if (entries[index].Enabled) return index;
            }
            return lastEnabledIndex >= 0 &&
                lastEnabledIndex < entries.Count &&
                entries[lastEnabledIndex].Enabled
                ? lastEnabledIndex
                : -1;
        }
    }

    internal enum SnippetMenuChoiceKind
    {
        Cancel,
        Insert,
        Manage
    }

    internal sealed class SnippetMenuChoice
    {
        private SnippetMenuChoice(SnippetMenuChoiceKind kind, Snippet snippet)
        {
            Kind = kind;
            Snippet = snippet;
        }

        public SnippetMenuChoiceKind Kind { get; private set; }
        public Snippet Snippet { get; private set; }

        public static SnippetMenuChoice Cancel()
        {
            return new SnippetMenuChoice(SnippetMenuChoiceKind.Cancel, null);
        }

        public static SnippetMenuChoice Insert(Snippet snippet)
        {
            if (snippet == null) throw new ArgumentNullException("snippet");
            return new SnippetMenuChoice(SnippetMenuChoiceKind.Insert, snippet);
        }

        public static SnippetMenuChoice Manage()
        {
            return new SnippetMenuChoice(SnippetMenuChoiceKind.Manage, null);
        }
    }

    internal static class SnippetMenuChoiceDispatcher
    {
        public static void Execute(
            SnippetMenuChoice choice,
            Action<Snippet> insert,
            Action manage)
        {
            if (choice == null) throw new ArgumentNullException("choice");
            if (insert == null) throw new ArgumentNullException("insert");
            if (manage == null) throw new ArgumentNullException("manage");
            if (choice.Kind == SnippetMenuChoiceKind.Insert)
            {
                insert(choice.Snippet);
            }
            else if (choice.Kind == SnippetMenuChoiceKind.Manage)
            {
                manage();
            }
        }
    }

    internal interface ISnippetMenuPopup : IDisposable
    {
        SnippetMenuChoice Choose(IWin32Window owner);
    }

    internal sealed class SnippetMenuPopup : Form, ISnippetMenuPopup
    {
        private const int ItemHeight = 27;
        private readonly ListBox list;
        private SnippetMenuChoice choice = SnippetMenuChoice.Cancel();
        private int lastEnabledIndex = -1;

        public SnippetMenuPopup(IList<Snippet> snippets)
        {
            if (snippets == null) throw new ArgumentNullException("snippets");
            FormBorderStyle = FormBorderStyle.FixedSingle;
            StartPosition = FormStartPosition.Manual;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;
            ShowInTaskbar = false;
            KeyPreview = true;

            IList<SnippetMenuEntry> entries = SnippetMenuEntry.Create(snippets);
            int visibleItems = Math.Min(entries.Count, 12);
            ClientSize = new Size(310, visibleItems * ItemHeight + 4);
            list = new ListBox
            {
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = ItemHeight,
                IntegralHeight = false,
                DataSource = entries
            };
            list.DrawItem += DrawEntry;
            list.MouseUp += ActivateMouseEntry;
            list.KeyDown += HandleListKey;
            list.SelectedIndexChanged += SkipDisabledEntry;
            Controls.Add(list);
            SelectFirstEnabledEntry();
        }

        public SnippetMenuChoice Choose(IWin32Window owner)
        {
            if (owner == null) throw new ArgumentNullException("owner");
            Location = SnippetPopupPlacement.Calculate(owner.Handle, Size);
            Log.Trace("ui.SnippetMenuPopup", "before");
            Show(owner);
            while (Visible)
            {
                Application.DoEvents();
                if (Visible) Thread.Sleep(10);
            }
            Log.Trace("ui.SnippetMenuPopup", "after-" + choice.Kind);
            return choice;
        }

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            if (Visible)
            {
                choice = SnippetMenuChoice.Cancel();
                Close();
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                choice = SnippetMenuChoice.Cancel();
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void DrawEntry(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var entry = (SnippetMenuEntry)list.Items[e.Index];
            e.DrawBackground();
            if (entry.Kind == SnippetMenuEntryKind.Separator)
            {
                int y = e.Bounds.Top + e.Bounds.Height / 2;
                using (var pen = new Pen(SystemColors.ControlDark))
                {
                    e.Graphics.DrawLine(
                        pen,
                        e.Bounds.Left + 8,
                        y,
                        e.Bounds.Right - 8,
                        y);
                }
                return;
            }

            Color color = entry.Enabled
                ? e.ForeColor
                : SystemColors.GrayText;
            TextRenderer.DrawText(
                e.Graphics,
                entry.Text,
                e.Font,
                new Rectangle(
                    e.Bounds.Left + 10,
                    e.Bounds.Top,
                    e.Bounds.Width - 20,
                    e.Bounds.Height),
                color,
                TextFormatFlags.Left |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }

        private void ActivateMouseEntry(object sender, MouseEventArgs e)
        {
            int index = list.IndexFromPoint(e.Location);
            if (index < 0) return;
            var entry = (SnippetMenuEntry)list.Items[index];
            if (!entry.Enabled) return;
            list.SelectedIndex = index;
            ActivateSelectedEntry();
        }

        private void HandleListKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            ActivateSelectedEntry();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private void ActivateSelectedEntry()
        {
            var entry = list.SelectedItem as SnippetMenuEntry;
            if (entry == null || !entry.Enabled) return;
            choice = entry.Kind == SnippetMenuEntryKind.Manage
                ? SnippetMenuChoice.Manage()
                : SnippetMenuChoice.Insert(entry.Snippet);
            Close();
        }

        private void SelectFirstEnabledEntry()
        {
            for (int index = 0; index < list.Items.Count; index++)
            {
                if (((SnippetMenuEntry)list.Items[index]).Enabled)
                {
                    list.SelectedIndex = index;
                    lastEnabledIndex = index;
                    return;
                }
            }
        }

        private void SkipDisabledEntry(object sender, EventArgs e)
        {
            var entry = list.SelectedItem as SnippetMenuEntry;
            if (entry == null) return;
            if (entry.Enabled)
            {
                lastEnabledIndex = list.SelectedIndex;
                return;
            }
            int selectedIndex = list.SelectedIndex;
            int replacement = SnippetMenuEntry.FindEnabledIndex(
                (IList<SnippetMenuEntry>)list.DataSource,
                selectedIndex,
                lastEnabledIndex);
            if (replacement >= 0)
            {
                list.SelectedIndex = replacement;
            }
        }
    }

    internal static class SnippetPopupPlacement
    {
        private const int RibbonOffset = 112;
        private const int RightMargin = 18;

        public static Point Calculate(IntPtr ownerHandle, Size popupSize)
        {
            Rectangle ownerBounds;
            NativeMethods.Rect rect;
            if (ownerHandle != IntPtr.Zero &&
                NativeMethods.GetWindowRect(ownerHandle, out rect))
            {
                ownerBounds = Rectangle.FromLTRB(
                    rect.Left,
                    rect.Top,
                    rect.Right,
                    rect.Bottom);
            }
            else
            {
                ownerBounds = Screen.PrimaryScreen.WorkingArea;
            }

            Rectangle workingArea = Screen.FromRectangle(ownerBounds).WorkingArea;
            return Calculate(ownerBounds, workingArea, popupSize);
        }

        internal static Point Calculate(
            Rectangle ownerBounds,
            Rectangle workingArea,
            Size popupSize)
        {
            int x = ownerBounds.Right - popupSize.Width - RightMargin;
            int y = ownerBounds.Top + RibbonOffset;
            x = Math.Max(workingArea.Left, Math.Min(x, workingArea.Right - popupSize.Width));
            y = Math.Max(workingArea.Top, Math.Min(y, workingArea.Bottom - popupSize.Height));
            return new Point(x, y);
        }

        private static class NativeMethods
        {
            [StructLayout(LayoutKind.Sequential)]
            internal struct Rect
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GetWindowRect(
                IntPtr windowHandle,
                out Rect rect);
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
