using System;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Threading;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EasyPlusAddressReply
{
    internal static class Localization
    {
        public static void ApplyOfficeUiCulture(Outlook.Application application)
        {
            if (application == null)
                return;

            try
            {
                int lcid = application.LanguageSettings.get_LanguageID(
                    Office.MsoAppLanguageID.msoLanguageIDUI);
                Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo(lcid);
            }
            catch
            {
                // Keep the process UI culture. ResourceManager will fall back to English.
            }
        }
    }

    internal static class Strings
    {
        private static readonly ResourceManager ResourceManager = new ResourceManager(
            "EasyPlusAddressReply.Resources.Strings",
            Assembly.GetExecutingAssembly());

        public static string Get(string key)
        {
            if (string.IsNullOrEmpty(key))
                return string.Empty;

            try
            {
                return ResourceManager.GetString(key, Thread.CurrentThread.CurrentUICulture) ?? key;
            }
            catch (MissingManifestResourceException)
            {
                return key;
            }
        }

        public static string Format(string key, params object[] args)
        {
            return string.Format(
                Thread.CurrentThread.CurrentUICulture,
                Get(key),
                args ?? Array.Empty<object>());
        }
    }
}
