using System;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EasyPlusAddressReply
{
    internal static class SenderIdentity
    {
        private const string PR_SENT_REPRESENTING_NAME_W =
            "http://schemas.microsoft.com/mapi/proptag/0x0042001F";
        private const string PR_SENT_REPRESENTING_ADDRTYPE_W =
            "http://schemas.microsoft.com/mapi/proptag/0x0064001F";
        private const string PR_SENT_REPRESENTING_EMAIL_ADDRESS_W =
            "http://schemas.microsoft.com/mapi/proptag/0x0065001F";
        private const string PR_SENT_REPRESENTING_SMTP_ADDRESS_W =
            "http://schemas.microsoft.com/mapi/proptag/0x5D02001F";

        private static readonly Regex PlusAddressRegex = new Regex(
            @"\A[a-z0-9.!#$%&'*+/=?^_`{|}~-]+\+[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,63}\z",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            TimeSpan.FromMilliseconds(50));

        public static void Apply(Outlook.MailItem reply, string smtpAddress)
        {
            if (reply == null)
                throw new ArgumentNullException(nameof(reply));

            smtpAddress = ValidatePlusSmtpAddress(smtpAddress);

            // Keep Outlook's existing account as the transport account; only change the represented
            // Internet sender identity. This mirrors choosing "From -> Other E-mail Address" manually.
            reply.SentOnBehalfOfName = smtpAddress;

            string[] schemas =
            {
                PR_SENT_REPRESENTING_ADDRTYPE_W,
                PR_SENT_REPRESENTING_EMAIL_ADDRESS_W,
                PR_SENT_REPRESENTING_SMTP_ADDRESS_W,
                PR_SENT_REPRESENTING_NAME_W
            };

            object[] values =
            {
                "SMTP",
                smtpAddress,
                smtpAddress,
                smtpAddress
            };

            // SetProperties does not throw on a per-property failure: it returns an array with one
            // entry per schema name describing the individual results. Ignoring it would let the
            // reply keep Outlook's original sender while the add-in reported success.
            ThrowOnPropertyErrors(reply.PropertyAccessor.SetProperties(schemas, values));
        }

        private static void ThrowOnPropertyErrors(object setPropertiesResult)
        {
            // A successful call returns null. Anything else is inspected for reported failures.
            if (!(setPropertiesResult is Array results))
                return;

            foreach (var result in results)
            {
                if (result is Exception error)
                    throw new InvalidOperationException("Setting the sender identity failed.", error);
            }
        }

        private static string ValidatePlusSmtpAddress(string smtpAddress)
        {
            if (string.IsNullOrWhiteSpace(smtpAddress))
                throw new ArgumentException("SMTP address is empty.", nameof(smtpAddress));

            string value = smtpAddress.Trim();

            // The address originates in an untrusted message. Accept only the same conservative
            // plus-address shape used by AliasDetector and reject all control/header syntax.
            if (value.Length > 320 || value.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0)
                throw new ArgumentException("SMTP address is invalid.", nameof(smtpAddress));

            int at = value.LastIndexOf('@');
            if (at <= 0 || at != value.IndexOf('@') || at > 64 || at >= value.Length - 1)
                throw new ArgumentException("SMTP address is invalid.", nameof(smtpAddress));

            // Ordinal comparisons only; see the matching note in AliasDetector.NormalizeAddress.
            string domain = value.Substring(at + 1);
            if (domain.Length > 253 ||
                domain[0] == '.' ||
                domain[domain.Length - 1] == '.' ||
                domain.IndexOf("..", StringComparison.Ordinal) >= 0)
            {
                throw new ArgumentException("SMTP address is invalid.", nameof(smtpAddress));
            }

            try
            {
                if (!PlusAddressRegex.IsMatch(value))
                    throw new ArgumentException("SMTP address is invalid.", nameof(smtpAddress));
            }
            catch (RegexMatchTimeoutException)
            {
                throw new ArgumentException("SMTP address is invalid.", nameof(smtpAddress));
            }

            return value;
        }
    }
}
