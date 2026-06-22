# Session Context

## Goal
Build and maintain a full-stack product ecosystem for **AKSA 10X FASTER** — professional ribbon add-ins for MS Word, Excel, Photoshop, and Illustrator. Includes: website, licensing, installer, admin tools, and GitHub Pages deployment.

## Constraints & Preferences
- Use bangla/english bilingual where appropriate
- Keep files self-contained; minimize external dependencies
- Website must be static (GitHub Pages)
- License key: 25-char uppercase alphanumeric, generated via server
- Installer: Inno Setup, self-signed, no xcopy/rmdir
- Default project root: `E:\Website\MS Office\MS Office\`
- Github site root: `E:\Website\MS Office\MS Office\Github\`
- NSIS installer project: `E:\Website\MS Office\MS Office\10x-faster.nsi`
- Installer output: `E:\Website\MS Office\MS Office\Setup\`
- Download placed at `E:\Website\MS Office\MS Office\Github\downloads\aksa-word-setup.exe`

## Progress — Done

### Installer (Inno Setup → NSIS)
- [x] Inno Setup script with custom branding, no xcopy/rmdir
- [x] Rebuilt with NSIS (`10x-faster.nsi`) — safer, smaller, no xcopy/no rmdir
- [x] Self-signed with `aksa-digital-point.pfx` using `SignTool`
- [x] Download available at `downloads/aksa-word-setup.exe`

### Website (`Github/`)
- [x] **index.html** — Full SaaS site: hero (slider + stats), features grid, products section, testimonials, pricing (4 products × 4 plans), FAQ accordion, CTA, footer. Pink-purple light theme.
- [x] **request-key.html** — Form: Name, Email, Phone, Product, License Type (trial/6mo/12mo/lifetime), Transaction ID. Sends via EmailJS + mailto fallback.
- [x] **purchase.html** — Payment info (bKash/Nagad/Rocket to 01670201266), same form as request-key with `[Purchase]` subject.
- [x] **send-key.html** — Admin page: type customer email, license key, duration, send via EmailJS.
- [x] **js/script.js** — Shared EmailJS logic: template params, subject formatting, toast notifications.
- [x] **css/style.css** — Pink-purple light gradient theme. CSS variables for theming. Glassmorphism, smooth animations, responsive.

### Licensing
- [x] Registry path: `HKLM\Software\AKSA 10X FASTER\License` (LicenseType, LicenseKey, Licensee, LicensedTo)
- [x] Trial tracking via `TrialStart` — persists across uninstall/reinstall
- [x] License types: trial, 6mo, 12mo, 24mo, lifetime

### Deployment
- [x] GitHub remote: `https://github.com/plabon02/Aksadigitalpoint.git`
- [x] Pushed to `main` branch
- [x] GitHub Pages enabled at `https://plabon02.github.io/Aksadigitalpoint/`

### PDF Add-in (`pdf-addin/`)
- [x] **Solution & Project** — `AksaPdfAddin.sln`, `AksaPdfAddin.csproj` (targets .NET Framework 4.8)
- [x] **PdfAddIn.cs** — COM add-in: Word→PDF, PDF→Word, Merge, Split, Extract, Protect, Unlock, Batch, PDF Info
- [x] **PdfService.cs** — Native Word Interop-based PDF operations
- [x] **AksaPdfAddin_Ribbon.xml** — 4 ribbon groups (Convert, Merge & Split, Protect, Info)
- [x] **AssemblyInfo.cs** — ComVisible, Guid, ProgId: `AksaPdfTools.Connect`

## Next Steps
- Verify GitHub Pages is serving the latest version
- Open `pdf-addin/AksaPdfAddin.sln` in Visual Studio 2022 and build
- Add NuGet packages (iTextSharp/iText7 for advanced PDF operations)
- Test EmailJS integration (request key + purchase + send key forms actually send email)
- Build NSIS installer for Excel, Photoshop, Illustrator add-ins
- Set up trial/license server API

## Critical Context
- **2025-09-12**: Initial website created with hero, features, products, pricing, FAQ, CTA
- **2025-09-29**: request-key.html created with EmailJS integration
- **2025-09-30**: installer rebuilt (NSIS), self-signed signed, GitHub Pages pushed
- **2025-10-01**: purchase.html created, send-key.html created, js/script.js shared logic
- Website redesigned multiple times: dark navy → indigo-teal → pink-purple → premium dark SaaS → pink-purple light (current)
- Theme toggle removed; staying with pink-purple light
- ScrollReveal loaded from CDN for scroll animations
- No `_config.yml` — GitHub Pages default Jekyll processing may affect paths

## Relevant Files
- `E:\Website\MS Office\MS Office\Github\index.html` — main site
- `E:\Website\MS Office\MS Office\Github\request-key.html` — license request form
- `E:\Website\MS Office\MS Office\Github\purchase.html` — purchase form
- `E:\Website\MS Office\MS Office\Github\send-key.html` — admin key sender
- `E:\Website\MS Office\MS Office\Github\css\style.css` — theme/styles
- `E:\Website\MS Office\MS Office\Github\js\script.js` — shared EmailJS logic
- `E:\Website\MS Office\MS Office\10x-faster.nsi` — NSIS installer script
- `E:\Website\MS Office\MS Office\aksa-digital-point.pfx` — code signing cert
- `E:\Website\MS Office\MS Office\Setup\aksa-word-setup.exe` — built installer
- `E:\Website\MS Office\MS Office\Github\downloads\aksa-word-setup.exe` — public download
- `E:\Website\MS Office\MS Office\pdf-addin\` — PDF add-in source code
