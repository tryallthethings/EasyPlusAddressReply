using Outlook = Microsoft.Office.Interop.Outlook;

namespace EasyPlusAddressReply
{
    internal sealed class ExplorerHook
    {
        private Outlook.Explorer _explorer;
        private Outlook.ExplorerEvents_10_Event _events;
        private readonly EasyPlusAddressReplyService _service;
        private readonly long _identity;

        public ExplorerHook(Outlook.Explorer explorer, long identity, EasyPlusAddressReplyService service)
        {
            _explorer = explorer;
            _identity = identity;
            _service = service;
            _events = (Outlook.ExplorerEvents_10_Event)explorer;
            _events.SelectionChange += OnSelectionChange;
            _events.Close += OnClose;
            OnSelectionChange();
        }

        private void OnSelectionChange()
        {
            _service.HookExplorerSelection(_explorer);
        }

        private void OnClose()
        {
            Detach();
            _service.ForgetExplorerHook(_identity);
        }

        public void Detach()
        {
            try
            {
                if (_events != null)
                {
                    _events.SelectionChange -= OnSelectionChange;
                    _events.Close -= OnClose;
                }
            }
            catch { }
            _explorer = null;
            _events = null;
        }
    }
}
