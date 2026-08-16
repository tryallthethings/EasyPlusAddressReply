# EasyPlusAddressReply

An add-in for classic Outlook on Windows that replies from the plus-address a message was actually sent to.

If mail arrives at `name+shop@example.com`, **Reply** and **Reply All** normally go out as `name@example.com` and the recipient loses track of which address they wrote to. EasyPlusAddressReply sets the From address back to `name+shop@example.com` automatically.

It needs an account whose SMTP server accepts the plus-address as a sender address — typically an ordinary IMAP/SMTP mailbox.

## Requirements

- Windows
- **Classic** Outlook for Windows (the traditional desktop application)
- .NET Framework 4.8
- Microsoft Visual Studio 2010 Tools for Office Runtime — the installer adds this if it is missing
- A mail account whose SMTP server permits sending from the plus-address

### Outlook versions

**Tested:** classic Outlook for Microsoft 365 (Office 2021 Pro Plus). This is the only configuration the add-in has been verified on.

**Untested, but expected to work:** Outlook 2013, 2016 and 2019. The ribbon uses the Office 2010 customUI schema and the project builds against the Office 2013 interop assemblies, so nothing in it should require a newer Outlook — but nobody has confirmed this. Reports welcome.

**Not supported:** the *new* Outlook for Windows, Outlook on the web, and Outlook for Mac. This is a VSTO add-in, and [Microsoft does not support VSTO or COM add-ins in new Outlook](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook).

## Installation

1. Close Outlook.
2. Download the release archive and extract it.
3. Run `setup.exe` from the extracted folder. Windows may ask you to confirm that you trust the publisher.
4. Start Outlook. An **EasyPlusAddressReply** tab appears in the ribbon.

Keep the extracted folder until the installation has finished. `setup.exe` is the recommended entry point because, unlike opening the `.vsto` file directly, it also installs the prerequisites.

The add-in does not check for updates. To update, download the newer archive and run its `setup.exe` again.

To uninstall, open the list of installed programs in Windows Settings or Control Panel, select **EasyPlusAddressReply** and choose Uninstall. Office add-ins cannot be removed from Outlook's own add-in list.

## Using it

The ribbon tab has three controls:

| Control | What it does |
|---|---|
| **Enabled / Disabled** | Master switch. When disabled, Outlook's normal From address is used. |
| **Settings** | Opens the settings window. |
| **Check message** | Explains which plus-address would be chosen for the message you have selected. |

Once enabled, there is nothing to do — replies pick up the right address on their own. Use **Check message** when a reply did not use the address you expected.

## How the address is chosen

There is no alias list to maintain. The add-in reads the SMTP address of each account configured in Outlook and accepts a candidate only if stripping its plus-tag yields one of those addresses:

```text
name+shop@example.com  →  name@example.com  ✓ matches a configured account
```

It looks first at Outlook's recipient list for the message. If that does not produce exactly one match, it can also read delivery headers such as `Delivered-To` and `X-Original-To`.

**It never guesses.** If more than one valid plus-address remains, the From address is left untouched. And because a candidate must reduce to an address that is already configured in Outlook, no incoming message can talk the add-in into sending from an unrelated address by putting one in a header.

## Settings

**Use delivery headers to resolve the receiving address** — when Outlook's recipient list is ambiguous, also inspect delivery headers. Enabled by default; recommended. Messages delivered via BCC often carry the receiving address only in these headers.

**Warn when multiple plus-addresses match** — show a warning when the add-in declines to choose. With this off, the add-in stays silent and simply leaves the From address alone.

The master on/off switch lives on the ribbon rather than being duplicated here.

The interface follows Outlook's own display language and is available in English and German.

## Privacy

Everything happens locally. The add-in does not send telemetry, contact any server, check for updates, or upload message data. It does not read message bodies.

The only things stored on disk are the two Boolean settings above, in `%LocalAppData%\EasyPlusAddressReply\settings.xml`. No addresses, headers or logs are written anywhere.

Message headers are processed in memory only, with a size limit and regular-expression timeouts, and the resulting address is validated independently before it is written to a reply. **Check message** masks addresses before displaying them and never shows raw headers.

## Troubleshooting

**A reply still used the base address.** Select the original message and use **Check message**. It reports, with addresses masked, whether the recipient list or the delivery headers produced a unique plus-address. The two common outcomes are that no candidate reduced to a configured account address, or that several candidates remained and the add-in declined to guess.

**The ribbon tab is missing.** Check that Outlook has not disabled the add-in: **File → Options → Add-ins**. Note that this is classic Outlook only — the tab will never appear in the new Outlook for Windows.

**The reply is rejected when sending.** The add-in only sets the From address; your SMTP server still has to accept the plus-address as a sender. Some providers do not.

## Building from source

Requires Visual Studio 2022 with the **Office/SharePoint development** workload, the .NET Framework 4.8 targeting pack, and desktop Outlook installed. Open the solution and build `Release`.

Note that the project signs its ClickOnce manifests with a specific certificate thumbprint. To build without that certificate, clear `ManifestCertificateThumbprint` in `EasyPlusAddressReply.csproj` or set `SignManifests` to `false`.

Keep `<ApplicationVersion>` in the `.csproj` in step with the version attributes in `Properties/AssemblyInfo.cs`; the settings window displays the latter.

## Technical notes

The project deliberately uses VSTO event handling for Reply and Reply All, RibbonX for the ribbon, WPF for the settings window, `.resx` files for localization, and a small local XML file for preferences.

RibbonX has no native switch control, so the ribbon uses Outlook's own `toggleButton` for the master on/off state; the settings window uses proper switch controls.

## License

MIT. See [LICENSE](LICENSE).
