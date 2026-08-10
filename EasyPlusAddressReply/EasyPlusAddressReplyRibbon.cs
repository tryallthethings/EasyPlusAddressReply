using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace EasyPlusAddressReply
{
    [ComVisible(true)]
    public sealed class EasyPlusAddressReplyRibbon : Office.IRibbonExtensibility
    {
        private static Office.IRibbonUI _ribbonUi;

        public string GetCustomUI(string ribbonID)
        {
            try { Localization.ApplyOfficeUiCulture(Globals.ThisAddIn.Application); } catch { }

            if (!string.Equals(ribbonID, "Microsoft.Outlook.Explorer", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(ribbonID, "Microsoft.Outlook.Mail.Read", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return @"<?xml version='1.0' encoding='UTF-8'?>
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='EparTab'
           insertAfterMso='TabAddIns'
           getLabel='GetLabel'>
        <group id='EparStatusGroup' getLabel='GetLabel'>
          <toggleButton id='EparEnabled'
                        size='large'
                        getLabel='GetLabel'
                        getImage='GetProductImage'
                        getPressed='GetPressed'
                        getScreentip='GetScreenTip'
                        onAction='OnToggleEnabled'/>
        </group>
        <group id='EparConfigurationGroup' getLabel='GetLabel'>
          <button id='EparSettings'
                  size='large'
                  imageMso='AdvancedFileProperties'
                  getLabel='GetLabel'
                  getScreentip='GetScreenTip'
                  onAction='OnSettings'/>
        </group>
        <group id='EparToolsGroup' getLabel='GetLabel'>
          <button id='EparDiagnose'
                  size='large'
                  imageMso='MessageHeaderToggle'
                  getLabel='GetLabel'
                  getScreentip='GetScreenTip'
                  onAction='OnDiagnose'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbonUi = ribbonUi;
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            switch (control?.Id)
            {
                case "EparTab": return Strings.Get("AppName");
                case "EparStatusGroup": return Strings.Get("RibbonStatusGroup");
                case "EparConfigurationGroup": return Strings.Get("RibbonConfigurationGroup");
                case "EparToolsGroup": return Strings.Get("RibbonToolsGroup");
                case "EparEnabled":
                    return GetEnabledState()
                        ? Strings.Get("RibbonEnabledOn")
                        : Strings.Get("RibbonEnabledOff");
                case "EparSettings": return Strings.Get("RibbonSettings");
                case "EparDiagnose": return Strings.Get("RibbonDiagnose");
                default: return Strings.Get("AppName");
            }
        }

        public string GetScreenTip(Office.IRibbonControl control)
        {
            switch (control?.Id)
            {
                case "EparEnabled": return Strings.Get("RibbonEnabledTip");
                case "EparSettings": return Strings.Get("RibbonSettingsTip");
                case "EparDiagnose": return Strings.Get("RibbonDiagnoseTip");
                default: return Strings.Get("AppName");
            }
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            return GetEnabledState();
        }

        public object GetProductImage(Office.IRibbonControl control)
        {
            return RibbonImages.GetProductIcon();
        }

        public void OnToggleEnabled(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.Service?.SetEnabled(pressed);
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Service?.ShowSettings();
        }

        public void OnDiagnose(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Service?.DiagnoseSelectedMessage();
        }

        internal static void ReleaseRibbon()
        {
            _ribbonUi = null;
            RibbonImages.Dispose();
        }

        internal static void InvalidateState()
        {
            try { _ribbonUi?.InvalidateControl("EparEnabled"); }
            catch { }
        }

        private static bool GetEnabledState()
        {
            return Globals.ThisAddIn.Service?.Settings.Enabled ?? true;
        }

        /// <summary>
        /// RibbonX has no native Windows-style switch. toggleButton is the native accessible
        /// on/off control. Outlook supplies the pressed/unpressed state; the product icon is
        /// intentionally static so the icon itself never has to communicate application state.
        /// </summary>
        private static class RibbonImages
        {
            private static Bitmap _productBitmap;
            private static stdole.IPictureDisp _productPicture;

            internal static object GetProductIcon()
            {
                if (_productPicture != null)
                    return _productPicture;

                try
                {
                    var resource = Properties.Resources.EasyPlusAddressReplyIcon;
                    if (resource == null)
                        return null;

                    // Clone the resource so this class exclusively owns the GDI object lifetime.
                    _productBitmap = new Bitmap(resource);
                    _productPicture = PictureDispConverter.ToPictureDisp(_productBitmap);
                    return _productPicture;
                }
                catch
                {
                    // A missing/corrupt cosmetic resource must never prevent the add-in from loading.
                    return null;
                }
            }

            internal static void Dispose()
            {
                if (_productPicture != null)
                {
                    try
                    {
                        if (Marshal.IsComObject(_productPicture))
                            Marshal.FinalReleaseComObject(_productPicture);
                    }
                    catch { }
                    finally
                    {
                        _productPicture = null;
                    }
                }

                _productBitmap?.Dispose();
                _productBitmap = null;
            }

            private sealed class PictureDispConverter : AxHost
            {
                private PictureDispConverter() : base(string.Empty) { }

                internal static stdole.IPictureDisp ToPictureDisp(Image image)
                {
                    return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
                }
            }
        }
    }
}
