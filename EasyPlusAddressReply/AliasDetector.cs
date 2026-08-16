using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EasyPlusAddressReply
{
    internal sealed class DetectionResult
    {
        public string SelectedAlias { get; set; }
        public List<string> RecipientCandidates { get; } = new List<string>();
        public List<string> HeaderCandidates { get; } = new List<string>();
        public List<string> BaseAddresses { get; } = new List<string>();
        public string Note { get; set; }

        public bool IsAmbiguous =>
            RecipientCandidates.Concat(HeaderCandidates)
                .Distinct(StringComparer.OrdinalIgnoreCase).Skip(1).Any();

        public string ToDiagnosticText()
        {
            var sb = new StringBuilder();
            sb.AppendLine(Strings.Get("DiagnosticTitle"));
            sb.AppendLine();
            sb.AppendLine(Strings.Get("DiagnosticBaseAddresses"));
            AppendRedactedAddresses(sb, BaseAddresses);
            sb.AppendLine();
            sb.AppendLine(Strings.Get("DiagnosticRecipientCandidates"));
            AppendRedactedAddresses(sb, RecipientCandidates);
            sb.AppendLine();
            sb.AppendLine(Strings.Get("DiagnosticHeaderCandidates"));
            AppendRedactedAddresses(sb, HeaderCandidates);
            sb.AppendLine();
            sb.AppendLine(Strings.Get("DiagnosticSelectedIdentity"));
            sb.AppendLine("  " + (SelectedAlias == null
                ? Strings.Get("DiagnosticNone")
                : PrivacyRedaction.RedactEmailAddress(SelectedAlias)));

            if (!string.IsNullOrWhiteSpace(Note))
            {
                sb.AppendLine();
                sb.AppendLine(Note);
            }

            sb.AppendLine();
            sb.AppendLine(Strings.Get("DiagnosticPrivacyNote"));
            return sb.ToString();
        }

        private static void AppendRedactedAddresses(StringBuilder sb, IEnumerable<string> addresses)
        {
            bool any = false;
            foreach (var address in addresses)
            {
                sb.AppendLine("  " + PrivacyRedaction.RedactEmailAddress(address));
                any = true;
            }

            if (!any)
                sb.AppendLine("  " + Strings.Get("DiagnosticNone"));
        }
    }

    internal sealed class AliasDetector
    {
        private readonly Outlook.Application _application;

        private const string PR_SMTP_ADDRESS_W =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";
        private const string PR_SMTP_ADDRESS_A =
            "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const string PR_TRANSPORT_MESSAGE_HEADERS_W =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001F";
        private const string PR_TRANSPORT_MESSAGE_HEADERS_A =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        // Deliberately fixed and ordered. Delivery/envelope headers are considered before visible To/Cc.
        // This is not user-configurable because guessing a different order can choose the wrong identity.
        private static readonly string[] HeaderNames =
        {
            "Delivered-To",
            "X-Original-To",
            "Envelope-To",
            "X-Envelope-To",
            "To",
            "Cc"
        };

        // Raw headers come from an untrusted message. Bound both input size and regex execution time.
        private const int MaxRawHeaderChars = 512 * 1024;
        private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(100);

        // Conservative Internet-address matcher for ordinary plus-addressing. The final address is
        // validated again by NormalizeAddress before it can become a sender identity.
        // Note: no inline (?x) option. The pattern contains '#' and must never be read as commented.
        private static readonly Regex PlusAddressRegex = new Regex(
            @"(?<![a-z0-9.!#$%&'*+/=?^_`{|}~-])" +
            @"([a-z0-9.!#$%&'*+/=?^_`{|}~-]{1,64}\+[a-z0-9.!#$%&'*+/=?^_`{|}~-]{1,64}@[a-z0-9.-]{1,253}\.[a-z]{2,63})" +
            @"(?![a-z0-9._%+-])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            RegexTimeout);

        public AliasDetector(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public DetectionResult Detect(Outlook.MailItem original, bool useDeliveryHeaderFallback)
        {
            var result = new DetectionResult();
            if (original == null)
            {
                result.Note = Strings.Get("DiagnosticNotMail");
                return result;
            }

            var bases = GetBaseAddresses().ToList();
            result.BaseAddresses.AddRange(bases);
            var baseSet = new HashSet<string>(bases, StringComparer.OrdinalIgnoreCase);
            if (baseSet.Count == 0)
            {
                result.Note = Strings.Get("DiagnosticNoAccountAddresses");
                return result;
            }

            var recipientMatches = FindRecipientAliases(original, baseSet)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            result.RecipientCandidates.AddRange(recipientMatches);

            if (recipientMatches.Count == 1)
            {
                result.SelectedAlias = recipientMatches[0];
                result.Note = Strings.Get("DiagnosticRecipientResolved");
                return result;
            }

            if (!useDeliveryHeaderFallback)
            {
                result.Note = recipientMatches.Count > 1
                    ? Strings.Get("DiagnosticAmbiguous")
                    : Strings.Get("DiagnosticNoMatch");
                return result;
            }

            var headerData = FindHeaderAliases(original, baseSet);
            result.HeaderCandidates.AddRange(headerData.AllAliases);
            var preferredHeaderAlias = headerData.PreferredAlias;

            if (recipientMatches.Count == 0 && !string.IsNullOrWhiteSpace(preferredHeaderAlias))
            {
                result.SelectedAlias = preferredHeaderAlias;
                result.Note = Strings.Get("DiagnosticHeaderResolved");
                return result;
            }

            if (recipientMatches.Count > 1 &&
                !string.IsNullOrWhiteSpace(preferredHeaderAlias) &&
                recipientMatches.Contains(preferredHeaderAlias, StringComparer.OrdinalIgnoreCase))
            {
                result.SelectedAlias = preferredHeaderAlias;
                result.Note = Strings.Get("DiagnosticAmbiguityResolved");
                return result;
            }

            result.Note = recipientMatches.Count > 1 || result.HeaderCandidates.Count > 1
                ? Strings.Get("DiagnosticAmbiguous")
                : Strings.Get("DiagnosticNoMatch");

            return result;
        }

        private IEnumerable<string> GetBaseAddresses()
        {
            var found = new List<string>();

            try
            {
                Outlook.Accounts accounts = _application.Session.Accounts;
                for (int i = 1; i <= accounts.Count; i++)
                {
                    try
                    {
                        Outlook.Account account = accounts[i];
                        var smtp = NormalizeAddress(account.SmtpAddress);
                        if (smtp != null)
                            found.Add(smtp);
                    }
                    catch
                    {
                        // One broken/unavailable account must not disable the add-in for all others.
                    }
                }
            }
            catch
            {
                // Outlook may transiently deny account enumeration during startup/shutdown.
            }

            return found.Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> FindRecipientAliases(
            Outlook.MailItem original,
            HashSet<string> baseSet)
        {
            Outlook.Recipients recipients;
            try
            {
                recipients = original.Recipients;
            }
            catch
            {
                yield break;
            }

            for (int i = 1; i <= recipients.Count; i++)
            {
                string address = null;
                try
                {
                    Outlook.Recipient recipient = recipients[i];
                    address = TryGetRecipientSmtpAddress(recipient);
                }
                catch
                {
                    // Skip an unreadable recipient rather than failing the whole reply.
                }

                if (address == null)
                    continue;

                foreach (var alias in ExtractValidAliases(address, baseSet))
                    yield return alias;
            }
        }

        private static string TryGetRecipientSmtpAddress(Outlook.Recipient recipient)
        {
            if (recipient == null)
                return null;

            try
            {
                var value = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS_W) as string;
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch { }

            try
            {
                var value = recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS_A) as string;
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch { }

            try
            {
                var value = recipient.Address;
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch { }

            try
            {
                Outlook.AddressEntry entry = recipient.AddressEntry;
                var value = entry?.Address;
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }
            catch { }

            return null;
        }

        private sealed class HeaderDetection
        {
            public List<string> AllAliases { get; } = new List<string>();
            public string PreferredAlias { get; set; }
        }

        private static HeaderDetection FindHeaderAliases(
            Outlook.MailItem original,
            HashSet<string> baseSet)
        {
            var result = new HeaderDetection();
            string rawHeaders = null;

            try
            {
                rawHeaders = original.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS_W) as string;
            }
            catch { }

            if (string.IsNullOrWhiteSpace(rawHeaders))
            {
                try
                {
                    rawHeaders = original.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS_A) as string;
                }
                catch { }
            }

            if (string.IsNullOrWhiteSpace(rawHeaders) || rawHeaders.Length > MaxRawHeaderChars)
                return result;

            var parsed = ParseHeaders(rawHeaders).ToList();
            foreach (var headerName in HeaderNames)
            {
                var aliasesForThisHeader = parsed
                    .Where(h => string.Equals(h.Key, headerName, StringComparison.OrdinalIgnoreCase))
                    .SelectMany(h => ExtractValidAliases(h.Value, baseSet))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var alias in aliasesForThisHeader)
                {
                    if (!result.AllAliases.Contains(alias, StringComparer.OrdinalIgnoreCase))
                        result.AllAliases.Add(alias);
                }

                if (result.PreferredAlias == null && aliasesForThisHeader.Count == 1)
                    result.PreferredAlias = aliasesForThisHeader[0];
            }

            return result;
        }

        private static IEnumerable<KeyValuePair<string, string>> ParseHeaders(string raw)
        {
            string currentName = null;
            var currentValue = new StringBuilder();

            foreach (var rawLine in raw.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            {
                // Ordinal by construction: header folding is defined in terms of raw bytes,
                // so culture-aware StringComparison must not take part in this decision.
                if (rawLine.Length > 0 && (rawLine[0] == ' ' || rawLine[0] == '\t') && currentName != null)
                {
                    currentValue.Append(' ').Append(rawLine.Trim());
                    continue;
                }

                if (currentName != null)
                    yield return new KeyValuePair<string, string>(currentName, currentValue.ToString());

                currentName = null;
                currentValue.Clear();

                int colon = rawLine.IndexOf(':');
                if (colon <= 0)
                    continue;

                currentName = rawLine.Substring(0, colon).Trim();
                currentValue.Append(rawLine.Substring(colon + 1).Trim());
            }

            if (currentName != null)
                yield return new KeyValuePair<string, string>(currentName, currentValue.ToString());
        }

        private static IEnumerable<string> ExtractValidAliases(string text, HashSet<string> baseSet)
        {
            if (string.IsNullOrWhiteSpace(text))
                return Enumerable.Empty<string>();

            var aliases = new List<string>();
            try
            {
                foreach (Match match in PlusAddressRegex.Matches(text))
                {
                    var alias = NormalizeAddress(match.Groups[1].Value);
                    if (alias == null)
                        continue;

                    var baseAddress = BaseOfPlusAddress(alias);
                    if (baseAddress != null && baseSet.Contains(baseAddress))
                        aliases.Add(alias);
                }
            }
            catch (RegexMatchTimeoutException)
            {
                // Treat a pathological header as no match. Never block Outlook on untrusted input.
            }

            return aliases;
        }

        private static string BaseOfPlusAddress(string address)
        {
            int at = address.LastIndexOf('@');
            if (at <= 0)
                return null;

            int plus = address.IndexOf('+');
            if (plus <= 0 || plus >= at - 1)
                return null;

            return address.Substring(0, plus) + address.Substring(at);
        }

        private static string NormalizeAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return null;

            address = address.Trim().Trim('<', '>', '"', '\'');
            if (address.Length > 320 || address.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0)
                return null;

            int at = address.LastIndexOf('@');
            if (at <= 0 || at != address.IndexOf('@') || at > 64 || at >= address.Length - 1)
                return null;

            // Ordinal comparisons only: this is a security check on untrusted input, and
            // culture-aware matching can ignore characters such as soft hyphens.
            string domain = address.Substring(at + 1);
            if (domain.Length > 253 ||
                domain[0] == '.' ||
                domain[domain.Length - 1] == '.' ||
                domain.IndexOf("..", StringComparison.Ordinal) >= 0)
            {
                return null;
            }

            return address.ToLowerInvariant();
        }
    }

    internal static class PrivacyRedaction
    {
        public static string RedactEmailAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return Strings.Get("DiagnosticNone");

            int at = address.LastIndexOf('@');
            if (at <= 0 || at >= address.Length - 1)
                return "***";

            string local = address.Substring(0, at);
            string domain = address.Substring(at + 1);
            int plus = local.IndexOf('+');

            string redactedLocal = plus > 0
                ? RedactToken(local.Substring(0, plus)) + "+" + RedactToken(local.Substring(plus + 1))
                : RedactToken(local);

            var labels = domain.Split('.');
            string redactedDomain;
            if (labels.Length >= 2)
            {
                labels[0] = RedactToken(labels[0]);
                redactedDomain = string.Join(".", labels);
            }
            else
            {
                redactedDomain = RedactToken(domain);
            }

            return redactedLocal + "@" + redactedDomain;
        }

        private static string RedactToken(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "***";
            return value.Length == 1 ? value + "***" : value.Substring(0, 1) + "***";
        }
    }
}
