# EasyPlusAddressReply

EasyPlusAddressReply is a small Outlook for Windows add-in that automatically selects the **plus-address that actually received a message** when you reply.

If mail arrives at:

```text
name+shop@example.com
```

EasyPlusAddressReply makes a normal **Reply** or **Reply All** use:

```text
From: name+shop@example.com
```

instead of falling back to:

```text
From: name@example.com
```

This is intended for ordinary IMAP/SMTP accounts where the SMTP server already accepts plus-addresses as sender addresses.

## Features

- Automatically selects the received plus-address for **Reply** and **Reply All**
- Uses the SMTP address of configured Outlook accounts as the trusted base identity
- Can use delivery headers such as `Delivered-To` and `X-Original-To` when Outlook's recipient list is ambiguous
- Never guesses when more than one valid plus-address remains
- Dedicated Outlook Ribbon tab
- English and German UI
- No telemetry, analytics, update checks, or external network calls
- Diagnostics redact email addresses and do not display or persist raw message headers

## Requirements

- Windows
- Classic Outlook for Windows (the traditional desktop Outlook application)
- Outlook 2021 or a Microsoft 365 desktop Outlook installation that supports VSTO add-ins
- .NET Framework 4.8
- Microsoft Visual Studio Tools for Office Runtime
- An IMAP/SMTP account whose SMTP server permits sending from the relevant plus-address

EasyPlusAddressReply is a VSTO add-in and does **not** run in the new Outlook for Windows.

## Settings

### Use delivery headers to resolve the receiving address

When Outlook's recipient collection does not identify one unique plus-address, EasyPlusAddressReply can inspect delivery headers such as `Delivered-To` and `X-Original-To`.

This is enabled by default and is recommended.

### Warn when multiple plus-addresses match

When more than one valid plus-address remains after all enabled checks, EasyPlusAddressReply can display a warning.

If the warning is disabled, the add-in remains silent and leaves Outlook's normal From address unchanged. It never chooses an ambiguous address.

The master **Enabled / Disabled** control is on the EasyPlusAddressReply Ribbon tab rather than being duplicated in the Settings window.

## How address selection works

EasyPlusAddressReply does not keep a manually configured alias list.

For each configured Outlook account, it obtains the account's SMTP address, for example:

```text
name@example.com
```

A candidate sender is accepted only when removing its plus-tag produces one of those configured account addresses:

```text
name+shop@example.com
        ↓
name@example.com
```

A message cannot therefore make the add-in send from an unrelated address simply by putting that address into a mail header.

## Privacy and security

EasyPlusAddressReply is designed to process mail data locally.

It does not:

- send telemetry
- contact an update server
- upload message data
- read message bodies
- persist raw message headers
- persist recipient or sender addresses
- write diagnostic logs containing email addresses

The only stored preferences are Boolean application settings.

Raw message headers, when needed, are processed in memory. Header input is size-limited, regular-expression matching uses execution timeouts, and sender addresses are independently validated before they are written to an Outlook reply.

The **Check message** diagnostic masks email addresses before displaying them.

## Building from source

### Prerequisites

Install:

- Visual Studio 2022
- the **Office/SharePoint development** workload
- .NET Framework 4.8 targeting tools
- desktop Outlook for Windows

Then open the project in Visual Studio and build the `Release` configuration.

The release output must include the VSTO deployment manifest (`.vsto`), application manifest (`.manifest`), the add-in assembly, and any generated satellite-resource directories such as `de`.

## Versioning

The product version is defined in:

```text
Properties/AssemblyInfo.cs
```

For a normal release, keep these values aligned:

```csharp
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: AssemblyInformationalVersion("1.0.0")]
```

The Settings window displays the informational version.

When creating an MSI, use the same three-part product version in the installer, for example `1.0.0`.

## Packaging with Advanced Installer

A ClickOnce deployment is not required.

Advanced Installer's dedicated **Office Add-In** packaging wizard requires its Professional edition, but a VSTO add-in can also be packaged manually as a normal Windows Installer package. Advanced Installer Freeware can create a Simple Installer project, install files, and create the registry values required by Outlook.

### Recommended x64 package

For 64-bit Outlook:

1. Build EasyPlusAddressReply in `Release`.
2. Create a **Simple Installer** project in Advanced Installer.
3. Install the complete required Release output to a directory under `Program Files`, for example:

   ```text
   [ProgramFiles64Folder]\EasyPlusAddressReply
   ```

4. Preserve generated subdirectories such as `de`.
5. Create this registry key in the **64-bit** registry view:

   ```text
   HKLM\Software\Microsoft\Office\Outlook\Addins\EasyPlusAddressReply
   ```

6. Add these values:

   | Name | Type | Value |
   |---|---|---|
   | `Description` | `REG_SZ` | `Automatically replies from the plus-address that received the message.` |
   | `FriendlyName` | `REG_SZ` | `EasyPlusAddressReply` |
   | `LoadBehavior` | `REG_DWORD` | `3` |
   | `Manifest` | `REG_SZ` | `file:///[APPDIR]EasyPlusAddressReply.vsto\|vstolocal` |

7. Mark the registry component as **64-bit** for a 64-bit Outlook package.
8. Require or document .NET Framework 4.8 and the Visual Studio Tools for Office Runtime.
9. Sign the finished MSI with your production code-signing certificate and a timestamp.
10. Test installation and removal on a clean Windows VM before publishing it.

Do not ship debug `.pdb` files unless you deliberately want to make symbols available.

### 32-bit Outlook

If you want to support both 32-bit and 64-bit Outlook while staying with the simple Freeware workflow, the clearest approach is to publish separate `x86` and `x64` MSI packages so the Outlook registration is written to the correct registry view.

The add-in itself is built as `AnyCPU`; the installer registration is the bitness-sensitive part.

### Upgrades

EasyPlusAddressReply has no built-in updater and does not need one.

For a future release, publish a new MSI with a higher version number and configure the Windows Installer project as a normal major upgrade. Keep the installer's Upgrade Code stable across versions and let the installer authoring tool manage the version-specific Product Code.

## Publishing on GitHub

A release can be kept simple:

1. Commit source code without build output or signing keys.
2. Tag the release, for example `v1.0.0`.
3. Create a GitHub Release from that tag.
4. Attach the signed installer:

   ```text
   EasyPlusAddressReply-1.0.0-x64.msi
   ```

5. If supported, also attach:

   ```text
   EasyPlusAddressReply-1.0.0-x86.msi
   ```

6. Include concise release notes and the SHA-256 hashes of the installer files.

Do **not** commit PFX/P12 files, certificate passwords, token PINs, private keys, or other signing credentials.

## Troubleshooting

### The UI does not match the source code

### Reply still uses the base address

Use **EasyPlusAddressReply → Check message** on the original message. The diagnostic shows, with addresses redacted, whether Outlook's recipient collection or the delivery headers produced a valid unique plus-address.

## Development notes

The project intentionally uses:

- VSTO/Outlook PIA event handling for Reply and Reply All
- RibbonX for the Outlook Ribbon
- WPF for the Settings UI
- `.resx` resources for localization
- local XML storage for the small Boolean preference set

RibbonX does not provide a native WinUI-style switch control. The Ribbon therefore uses Outlook's native `toggleButton` control for the master on/off state. The Settings window uses actual switch-style WPF toggle controls.

## License

Add a `LICENSE` file before publishing the repository publicly. Choose the license that matches how you want others to use, modify, and redistribute the project.


## Technical references

- [Microsoft: Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins)
- [Microsoft: Deploying a VSTO solution using Windows Installer](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer)
- [Microsoft: VSTO Add-in architecture](https://learn.microsoft.com/en-us/visualstudio/vsto/architecture-of-vsto-add-ins)
- [Advanced Installer: Office VSTO Add-in deployment](https://www.advancedinstaller.com/office-addin-deployment.html)
