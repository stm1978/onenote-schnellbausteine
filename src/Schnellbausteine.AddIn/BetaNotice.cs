using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Schnellbausteine.AddIn
{
    internal interface IBetaNoticePresenter
    {
        void Show(BetaStatus status, IntPtr ownerHandle);
    }

    internal interface IWebsiteLauncher
    {
        void Open(Uri website);
    }

    internal sealed class ShellWebsiteLauncher : IWebsiteLauncher
    {
        public void Open(Uri website)
        {
            Process.Start(new ProcessStartInfo(website.AbsoluteUri)
            {
                UseShellExecute = true
            });
        }
    }

    internal static class BetaWebsite
    {
        public static readonly Uri DownloadUri = new Uri("https://www.malter365.de/");

        public static void Open(
            IWebsiteLauncher launcher,
            Action<Exception> showError)
        {
            if (launcher == null) throw new ArgumentNullException("launcher");
            if (showError == null) throw new ArgumentNullException("showError");
            try
            {
                launcher.Open(DownloadUri);
            }
            catch (Exception exception)
            {
                showError(exception);
            }
        }
    }

    internal sealed class WinFormsBetaNoticePresenter : IBetaNoticePresenter
    {
        private readonly IWebsiteLauncher websiteLauncher;

        public WinFormsBetaNoticePresenter(IWebsiteLauncher websiteLauncher)
        {
            if (websiteLauncher == null) throw new ArgumentNullException("websiteLauncher");
            this.websiteLauncher = websiteLauncher;
        }

        public void Show(BetaStatus status, IntPtr ownerHandle)
        {
            using (var dialog = new BetaNoticeForm(status, websiteLauncher))
            {
                if (ownerHandle == IntPtr.Zero)
                {
                    dialog.ShowDialog();
                }
                else
                {
                    dialog.ShowDialog(new WindowHandleOwner(ownerHandle));
                }
            }
        }
    }

    internal sealed class BetaNoticeForm : Form
    {
        private readonly IWebsiteLauncher websiteLauncher;

        public BetaNoticeForm(BetaStatus status, IWebsiteLauncher websiteLauncher)
        {
            if (websiteLauncher == null) throw new ArgumentNullException("websiteLauncher");
            this.websiteLauncher = websiteLauncher;

            Text = "OneNote Schnellbausteine – Beta 0.2";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(500, 155);

            string message = status == BetaStatus.Expired
                ? "Diese Beta-Version von OneNote Schnellbausteine ist abgelaufen. " +
                  "Laden Sie die aktuelle Version herunter:"
                : "Diese Beta-Version von OneNote Schnellbausteine kann bis zum " +
                  "31. Dezember 2026 verwendet werden. Die aktuelle Version finden Sie auf:";
            var messageLabel = new Label
            {
                Left = 14,
                Top = 14,
                Width = 470,
                Height = 55,
                Text = message
            };
            var link = new LinkLabel
            {
                Left = 14,
                Top = 72,
                Width = 470,
                Height = 24,
                Text = BetaWebsite.DownloadUri.AbsoluteUri
            };
            link.LinkClicked += OpenWebsite;
            var close = new Button
            {
                Left = 399,
                Top = 112,
                Width = 85,
                Text = "Schließen",
                DialogResult = DialogResult.OK
            };
            Controls.AddRange(new Control[] { messageLabel, link, close });
            AcceptButton = close;
            CancelButton = close;
        }

        private void OpenWebsite(object sender, LinkLabelLinkClickedEventArgs e)
        {
            BetaWebsite.Open(
                websiteLauncher,
                delegate(Exception exception)
                {
                    MessageBox.Show(
                        this,
                        "Die Website konnte nicht geöffnet werden.\n\n" + exception.Message,
                        "OneNote Schnellbausteine – Beta 0.2",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                });
        }
    }
}
