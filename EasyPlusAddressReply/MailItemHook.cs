using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EasyPlusAddressReply
{
    internal sealed class MailItemHook
    {
        private Outlook.MailItem _item;
        private Outlook.ItemEvents_10_Event _events;
        private readonly EasyPlusAddressReplyService _service;
        private readonly long _identity;

        public MailItemHook(Outlook.MailItem item, long identity, EasyPlusAddressReplyService service)
        {
            _item = item ?? throw new ArgumentNullException(nameof(item));
            _identity = identity;
            _service = service ?? throw new ArgumentNullException(nameof(service));

            // MailItem has both Reply()/ReplyAll() methods and Reply/ReplyAll events.
            // Microsoft explicitly requires the latest events interface to connect to the events.
            _events = (Outlook.ItemEvents_10_Event)item;
            _events.Reply += OnReply;
            _events.ReplyAll += OnReplyAll;
            _events.Unload += OnUnload;
        }

        private void OnReply(object response, ref bool cancel)
        {
            _service.HandleReply(_item, response as Outlook.MailItem);
        }

        private void OnReplyAll(object response, ref bool cancel)
        {
            _service.HandleReply(_item, response as Outlook.MailItem);
        }

        private void OnUnload()
        {
            // Microsoft documents Unload as the point where add-ins should dereference the item.
            // Do not call methods/properties on the MailItem here.
            var service = _service;
            _item = null;
            _events = null;
            service.ForgetMailHook(_identity);
        }

        public void Detach()
        {
            try
            {
                if (_events != null)
                {
                    _events.Reply -= OnReply;
                    _events.ReplyAll -= OnReplyAll;
                    _events.Unload -= OnUnload;
                }
            }
            catch { }
            _item = null;
            _events = null;
        }
    }
}
