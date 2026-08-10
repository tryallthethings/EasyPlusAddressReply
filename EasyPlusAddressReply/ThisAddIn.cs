using System;
using Office = Microsoft.Office.Core;

namespace EasyPlusAddressReply
{
    public partial class ThisAddIn
    {
        internal EasyPlusAddressReplyService Service { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Localization.ApplyOfficeUiCulture(Application);

            Service = new EasyPlusAddressReplyService(Application);
            Service.Start();
            EasyPlusAddressReplyRibbon.InvalidateState();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Service?.Dispose();
            Service = null;
            EasyPlusAddressReplyRibbon.ReleaseRibbon();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new EasyPlusAddressReplyRibbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
