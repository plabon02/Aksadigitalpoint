# Session Context

## Goal
Build and maintain a full-stack product ecosystem for **AKSA 10X FASTER** — professional ribbon add-ins for MS Word, Excel, Photoshop, and Illustrator. Includes: website, licensing, installer, admin tools, and GitHub Pages deployment.

## Constraints & Preferences
- Use bangla/english bilingual where appropriate
- Keep files self-contained; minimize external dependencies
- Website must be static (GitHub Pages)
- License key: 25-char uppercase alphanumeric, generated via server
- NSIS installer: `10x-faster.nsi`
- Installer output: `E:\Website\MS Office\MS Office\Setup\`
- All website files live at repo root (moved from `Github/` to fix 404)
- GitHub Pages serves from root of `main` branch

## Progress — Done

### Website (Repo Root)
- [x] **index.html** — Full SaaS site: hero (slider + stats), features grid, products section, testimonials, pricing (4 products × 4 plans), FAQ accordion, CTA, footer. Pink-purple light theme.
- [x] **request-key.html** — Form: Name, Email, Phone, Product, License Type (trial/6mo/12mo/lifetime), Transaction ID. Sends via EmailJS + mailto fallback.
- [x] **purchase.html** — Payment info (bKash/Nagad/Rocket to 01670201266), same form as request-key with `[Purchase]` subject.
- [x] **send-key.html** — Admin page: type customer email, license key, duration, send via EmailJS.
- [x] **js/script.js** — Shared EmailJS logic: template params, subject formatting, toast notifications.
- [x] **css/style.css** — Pink-purple light gradient theme. CSS variables for theming. Glassmorphism, smooth animations, responsive.

### Installer (Inno Setup → NSIS)
- [x] NSIS installer (`10x-faster.nsi`) — safe, small, no xcopy/rmdir
- [x] Self-signed with `aksa-digital-point.pfx` using `SignTool`
- [x] Word add-in: `downloads/aksa-word-setup.exe`
- [x] Photoshop add-in 4 versions: `downloads/aksa-photoshop-setup-v1.00.exe` through `v1.40.exe`

### PDF Add-in (`pdf-addin/`)
- [x] **Solution & Project** — `AksaPdfAddin.sln`, `AksaPdfAddin.csproj` (targets .NET Framework 4.8)
- [x] **PdfAddIn.cs** — COM add-in with 7+ features including Open PDF (opens in Word)
- [x] **PdfService.cs** — Native Word Interop-based PDF operations (10 methods)
- [x] **AksaPdfAddin_Ribbon.xml** — 7 groups (File, Convert, Organize, Optimize, Edit, Security, Intelligence)
- [x] **AssemblyInfo.cs** — ComVisible, Guid, ProgId: `AksaPdfTools.Connect`
- [x] **Build succeeds** — Release + Debug, 0 errors, 0 warnings
- [x] **Output** — `pdf-addin\src\AksaPdfAddin\bin\Release\AksaPdfAddin.dll`
- [x] **Registered** — HKCU-based `register-pdf.ps1` script, no admin needed

### Licensing
- [x] Registry path: `HKLM\Software\AKSA 10X FASTER\License` (LicenseType, LicenseKey, Licensee, LicensedTo)
- [x] Trial tracking via `TrialStart` — persists across uninstall/reinstall
- [x] License types: trial, 6mo, 12mo, 24mo, lifetime

### Deployment
- [x] GitHub remote: `https://github.com/plabon02/Aksadigitalpoint.git`
- [x] Pushed to `main` branch
- [x] GitHub Pages at `https://plabon02.github.io/Aksadigitalpoint/` — serves from root
- [x] Files moved from `Github/` to repo root (fixes 404)

## Next Steps
- Test PDF add-in in Word (restart Word, check ribbon)
- Add PDF download links to website for Photoshop add-in versions
- Build NSIS installer for Excel, Photoshop, Illustrator add-ins
- Set up trial/license server API
- Test EmailJS integration

## Critical Context
- **2025-09-12**: Initial website created with hero, features, products, pricing, FAQ, CTA
- **2025-09-29**: request-key.html created with EmailJS integration
- **2025-09-30**: installer rebuilt (NSIS), self-signed signed, GitHub Pages pushed
- **2025-10-01**: purchase.html created, send-key.html created, js/script.js shared logic
- **2025-10-02**: PDF add-in created, built, registered; GitHub/ moved to root (404 fix)
- Website redesigned multiple times: dark navy → indigo-teal → pink-purple → premium dark SaaS → pink-purple light (current)
- ScrollReveal loaded from CDN for scroll animations

## Relevant Files
- `E:\Website\MS Office\MS Office\index.html` — main site
- `E:\Website\MS Office\MS Office\request-key.html` — license request form
- `E:\Website\MS Office\MS Office\purchase.html` — purchase form
- `E:\Website\MS Office\MS Office\send-key.html` — admin key sender
- `E:\Website\MS Office\MS Office\css\style.css` — theme/styles
- `E:\Website\MS Office\MS Office\js\script.js` — shared EmailJS logic
- `E:\Website\MS Office\MS Office\downloads\` — installer downloads
- `E:\Website\MS Office\MS Office\10x-faster.nsi` — NSIS installer script
- `E:\Website\MS Office\MS Office\pdf-addin\` — PDF add-in source code
- `E:\Website\MS Office\MS Office\pdf-addin\register-pdf.ps1` — registration script
