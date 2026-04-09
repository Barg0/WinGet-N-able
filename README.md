# 📦 WinGet N-able

PowerShell scripts for **one WinGet package per policy** (for example N-able N-sight RMM). Each run targets a single `winget` package id. 🙂

## ✅ Requirements

- Windows with **Windows Package Manager (winget)** available (system context preferred).
- **PowerShell 5.1 or later**, elevated if your deployment installs machine-wide.

## 🧰 Scripts

| Script | Role |
|--------|------|
| `detection.ps1` | 🔎 Reports whether the device needs remediation for the given package (missing install or upgrade pending). |
| `install.ps1` | ⬇️ Installs the package if missing, or upgrades it when an update is available. |

## 🎛️ Parameters

### `detection.ps1`

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-AppId` | Yes | WinGet package id (e.g. `Mozilla.Firefox`). |

**Exit codes**

- **0** ✅ — Compliant (installed and no pending upgrade from WinGet), or WinGet is unusable and the script exits without treating that as a hard failure (defer / retry in your platform).
- **1** ⚠️ — Non-compliant: package not installed, pending upgrade, resolve/list errors, or other detection failure.

### `install.ps1`

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-AppId` | Yes | WinGet package id. |
| `-Override` | No | Passed to `winget install` / upgrade flow as `--override` for silent installer switches. |

**Exit codes**

- **0** ✅ — Success, or a **transient / defer** outcome (e.g. installer busy, WinGet unhealthy) so automation can retry later without marking a hard failure.
- **1** ❌ — Install or upgrade failed after retries.

Installs use **`--scope machine`** first when applicable; scope / source retries are built in. Optional behavior is controlled at the top of `install.ps1` (locale workaround, wait/retry counts, and flags such as `$wingetUseInstallVersionFallback` and `$wingetUseUninstallPrevious`).

## 📝 Logs

Logs are written under:

`%ProgramData%\WinGetNable\<sanitized AppId>\`

- `detection.log` — output from `detection.ps1`
- `install.log` — output from `install.ps1`

The folder name is derived from `AppId` with characters unsafe for paths replaced.

## 💡 Examples

Run detection for a package:

```powershell
.\detection.ps1 -AppId 'Google.Chrome'
```

Install or upgrade the same package:

```powershell
.\install.ps1 -AppId 'Google.Chrome'
```

Pass an installer override (silent flags for the vendor EXE/MSI):

```powershell
.\install.ps1 -AppId 'Vendor.Product' -Override '/quiet /norestart'
```

From a scheduled task or RMM script action, call either script with the same `-AppId` your policy is scoped to. 🚀
