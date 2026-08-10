using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EasyPlusAddressReply
{
    internal sealed class EasyPlusAddressReplyService : IDisposable
    {
        private readonly Outlook.Application _application;
        private readonly Outlook.ApplicationEvents_11_Event _applicationEvents;
        private readonly Outlook.Explorers _explorers;
        private readonly Outlook.ExplorersEvents_Event _explorersEvents;
        private readonly SettingsStore _store = new SettingsStore();
        private readonly AliasDetector _detector;
        private readonly Dictionary<long, MailItemHook> _mailHooks = new Dictionary<long, MailItemHook>();
        private readonly Dictionary<long, ExplorerHook> _explorerHooks = new Dictionary<long, ExplorerHook>();
        private bool _disposed;

        public EasyPlusAddressReplySettings Settings { get; private set; }

        public EasyPlusAddressReplyService(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _applicationEvents = (Outlook.ApplicationEvents_11_Event)_application;
            _explorers = _application.Explorers;
            _explorersEvents = (Outlook.ExplorersEvents_Event)_explorers;
            _detector = new AliasDetector(_application);
            Settings = _store.Load();
        }

        public void Start()
        {
            _applicationEvents.ItemLoad += OnItemLoad;
            _explorersEvents.NewExplorer += OnNewExplorer;

            try
            {
                for (int i = 1; i <= _explorers.Count; i++)
                    TryHookExplorer(_explorers[i]);
            }
            catch { }

            try
            {
                var inspector = _application.ActiveInspector();
                if (inspector != null)
                    TryHookMailItem(inspector.CurrentItem);
            }
            catch { }
        }

        private void OnItemLoad(object item)
        {
            // ItemLoad fires before most item properties are available. Attach only the event sink here.
            TryHookMailItem(item);
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            TryHookExplorer(explorer);
        }

        internal void HookExplorerSelection(Outlook.Explorer explorer)
        {
            if (explorer == null)
                return;

            try
            {
                Outlook.Selection selection = explorer.Selection;
                for (int i = 1; i <= selection.Count; i++)
                    TryHookMailItem(selection[i]);
            }
            catch { }
        }

        private void TryHookMailItem(object item)
        {
            if (_disposed || !(item is Outlook.MailItem mail))
                return;

            long id;
            try { id = ComIdentity(mail); }
            catch { return; }

            if (_mailHooks.ContainsKey(id))
                return;

            try { _mailHooks[id] = new MailItemHook(mail, id, this); }
            catch { }
        }

        private void TryHookExplorer(Outlook.Explorer explorer)
        {
            if (_disposed || explorer == null)
                return;

            long id;
            try { id = ComIdentity(explorer); }
            catch { return; }

            if (_explorerHooks.ContainsKey(id))
                return;

            try { _explorerHooks[id] = new ExplorerHook(explorer, id, this); }
            catch { }
        }

        internal void ForgetMailHook(long identity)
        {
            _mailHooks.Remove(identity);
        }

        internal void ForgetExplorerHook(long identity)
        {
            _explorerHooks.Remove(identity);
        }

        internal void HandleReply(Outlook.MailItem original, Outlook.MailItem reply)
        {
            if (_disposed || !Settings.Enabled || original == null || reply == null)
                return;

            try
            {
                var result = _detector.Detect(original, Settings.UseDeliveryHeaderFallback);
                if (!string.IsNullOrWhiteSpace(result.SelectedAlias))
                {
                    SenderIdentity.Apply(reply, result.SelectedAlias);
                    return;
                }

                if (result.IsAmbiguous && Settings.WarnOnAmbiguousMatch)
                {
                    MessageBox.Show(
                        Strings.Get("MessageAmbiguous"),
                        Strings.Get("MessageBoxTitle"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch
            {
                // Do not display exception details: Outlook/MAPI exceptions can include account/folder data.
                MessageBox.Show(
                    Strings.Get("MessageApplyFailed"),
                    Strings.Get("MessageBoxTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        public void SetEnabled(bool enabled)
        {
            if (_disposed || Settings.Enabled == enabled)
                return;

            var updated = Settings.Clone();
            updated.Enabled = enabled;
            SaveSettings(updated);
        }

        public void ShowSettings()
        {
            if (_disposed)
                return;

            Localization.ApplyOfficeUiCulture(_application);

            var window = new OptionsWindow(
                Settings.UseDeliveryHeaderFallback,
                Settings.WarnOnAmbiguousMatch);

            if (window.ShowDialogOwned() != true)
                return;

            var updated = Settings.Clone();
            updated.UseDeliveryHeaderFallback = window.UseDeliveryHeaderFallback;
            updated.WarnOnAmbiguousMatch = window.WarnOnAmbiguousMatch;
            SaveSettings(updated);
        }

        public void DiagnoseSelectedMessage()
        {
            Localization.ApplyOfficeUiCulture(_application);

            try
            {
                Outlook.MailItem mail = GetActiveMailItem();
                if (mail == null)
                {
                    MessageBox.Show(
                        Strings.Get("MessageSelectOne"),
                        Strings.Get("MessageBoxTitle"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                var result = _detector.Detect(mail, Settings.UseDeliveryHeaderFallback);
                MessageBox.Show(
                    result.ToDiagnosticText(),
                    Strings.Get("DiagnosticTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch
            {
                MessageBox.Show(
                    Strings.Get("MessageDiagnosisFailed"),
                    Strings.Get("MessageBoxTitle"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SaveSettings(EasyPlusAddressReplySettings updated)
        {
            Settings = updated ?? throw new ArgumentNullException(nameof(updated));

            try { _store.Save(Settings); }
            catch
            {
                // Keep the in-memory setting for this Outlook session if persistence fails.
            }

            EasyPlusAddressReplyRibbon.InvalidateState();
        }

        private Outlook.MailItem GetActiveMailItem()
        {
            object activeWindow = null;
            try { activeWindow = _application.ActiveWindow(); }
            catch { }

            try
            {
                if (activeWindow is Outlook.Inspector inspector)
                    return inspector.CurrentItem as Outlook.MailItem;

                if (activeWindow is Outlook.Explorer explorer)
                {
                    Outlook.Selection selection = explorer.Selection;
                    if (selection.Count == 1)
                        return selection[1] as Outlook.MailItem;
                }
            }
            catch { }

            return null;
        }

        private static long ComIdentity(object comObject)
        {
            IntPtr unknown = Marshal.GetIUnknownForObject(comObject);
            try { return unknown.ToInt64(); }
            finally { Marshal.Release(unknown); }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;

            try { _applicationEvents.ItemLoad -= OnItemLoad; } catch { }
            try { _explorersEvents.NewExplorer -= OnNewExplorer; } catch { }

            foreach (var hook in _mailHooks.Values)
                hook.Detach();
            foreach (var hook in _explorerHooks.Values)
                hook.Detach();

            _mailHooks.Clear();
            _explorerHooks.Clear();
        }
    }
}
