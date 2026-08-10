using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Drawing.Imaging;

namespace EasyPlusAddressReply
{
    public partial class OptionsWindow : Window
    {
        public bool UseDeliveryHeaderFallback => UseDeliveryHeadersToggle.IsChecked == true;
        public bool WarnOnAmbiguousMatch => WarnOnAmbiguousToggle.IsChecked == true;

        public OptionsWindow(bool useDeliveryHeaderFallback, bool warnOnAmbiguousMatch)
        {
            InitializeComponent();

            ApplyProductIcon();

            Title = Strings.Get("SettingsTitle");
            HeadingText.Text = Strings.Get("SettingsHeading");
            UseDeliveryHeadersTitleText.Text = Strings.Get("SettingsUseDeliveryHeaders");
            UseDeliveryHeadersDescriptionText.Text = Strings.Get("SettingsUseDeliveryHeadersDescription");
            WarnOnAmbiguousTitleText.Text = Strings.Get("SettingsWarnAmbiguous");
            WarnOnAmbiguousDescriptionText.Text = Strings.Get("SettingsWarnAmbiguousDescription");
            VersionText.Text = Strings.Format("SettingsVersion", GetDisplayVersion());
            OkButton.Content = Strings.Get("OK");
            CancelButton.Content = Strings.Get("Cancel");

            UseDeliveryHeadersToggle.IsChecked = useDeliveryHeaderFallback;
            WarnOnAmbiguousToggle.IsChecked = warnOnAmbiguousMatch;
            UpdateStateLabels();
        }

        public bool? ShowDialogOwned()
        {
            try
            {
                IntPtr owner = NativeMethods.GetActiveWindow();
                if (owner != IntPtr.Zero)
                    new WindowInteropHelper(this).Owner = owner;
            }
            catch
            {
                // Falling back to a normal modal WPF window is safe if an HWND cannot be resolved.
            }

            return ShowDialog();
        }

        private void SettingToggle_StateChanged(object sender, RoutedEventArgs e)
        {
            UpdateStateLabels();
        }

        private void UpdateStateLabels()
        {
            if (UseDeliveryHeadersStateText != null)
            {
                UseDeliveryHeadersStateText.Text = UseDeliveryHeadersToggle.IsChecked == true
                    ? Strings.Get("StateOn")
                    : Strings.Get("StateOff");
            }

            if (WarnOnAmbiguousStateText != null)
            {
                WarnOnAmbiguousStateText.Text = WarnOnAmbiguousToggle.IsChecked == true
                    ? Strings.Get("StateOn")
                    : Strings.Get("StateOff");
            }
        }

        private static string GetDisplayVersion()
        {
            try
            {
                var assembly = typeof(OptionsWindow).Assembly;
                var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
                if (!string.IsNullOrWhiteSpace(informational))
                    return informational;

                var version = assembly.GetName().Version;
                if (version != null)
                    return $"{version.Major}.{version.Minor}.{Math.Max(version.Build, 0)}";
            }
            catch { }

            return "1.0.0";
        }

        private void ApplyProductIcon()
        {
            try
            {
                var bitmap = Properties.Resources.EasyPlusAddressReplyIcon;
                if (bitmap == null)
                    return;

                using (var stream = new MemoryStream())
                {
                    bitmap.Save(stream, ImageFormat.Png);
                    stream.Position = 0;

                    var image = new BitmapImage();
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = stream;
                    image.EndInit();
                    image.Freeze();

                    Icon = image;
                }
            }
            catch
            {
                // Cosmetic failure only; use Outlook/Windows' default icon.
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll")]
            internal static extern IntPtr GetActiveWindow();
        }
    }
}
