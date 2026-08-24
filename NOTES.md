# Project Progress & Next Steps

**Repository**: `Fralin-Development/fralin-development.github.io`  
**Current Branch**: `dev_steph_soph`  
**Date**: August 21, 2026

---

## 1. What We Completed Today

### Git & GitHub Configuration
- [x] **Global Git Setup**: Configured global `user.name` (`st33fo`) and `user.email` (`stefanpf96@gmail.com`).
- [x] **Default Branch**: Configured `init.defaultBranch = main`.
- [x] **SSH Connection**: Verified authentication with GitHub over SSH via `~/.ssh/id_ed25519`.
- [x] **Remote URL**: Switched remote origin from HTTPS to SSH (`git@github.com:Fralin-Development/fralin-development.github.io.git`).
- [x] **Branch Alignment**: Renamed local branch from `dev_stef_soph` to `dev_steph_soph` and linked upstream tracking to `origin/dev_steph_soph`.

### Technical Architecture Clarification
- [x] Researched and established the keyless method to fetch live Google Sheets data via CSV export (`/export?format=csv&gid={GID}`).
- [x] Identified support for targeting specific sheet tabs via `gid` or tab names without requiring Google Cloud Console or API keys.

---

## 2. Implementation Progress & Next Steps

- [x] **Commit Pending Files**: `_config.yml` committed and tracked on `dev_steph_soph`.
- [x] **Update `script.js`**:
  - [x] Implemented auto-parser `getGoogleSheetCsvUrl()` for Google Sheets and `getSharePointDownloadUrl()` for SharePoint / Excel Online.
  - [x] Added SheetJS (`xlsx.full.min.js`) support to read `.xlsx` workbooks and target specific tabs (e.g. `Donation Screen`).
  - [x] Implemented direct, keyless CSV fetching and robust RFC 4180 parsing (`parseCSV()`).
  - [x] Added smart column detection for Donor Name, Amount, and Custom Dedication/Message.
  - [x] Added local fallback (`donors.csv`) for resilience against external CORS or network restrictions.
  - [x] Configured auto-polling with cache-busting timestamp (`_t=...`) to reflect live updates automatically.
  - [x] Implemented dynamic animation duration calculation based on content height and configurable scroll speed.
  - [x] Added sample fallback donor roll and user-friendly status banners.
- [x] **Customize Scrolling & Visuals (`style.css` & `index.html`)**:
  - [x] Designed elegant museum/kiosk aesthetic with ambient radial lighting and gold accents.
  - [x] Configured smooth top & bottom edge masking (`mask-image: linear-gradient`).
  - [x] Added pause-on-hover capability and responsive sizing for mobile, 1080p, and 4K displays.
  - [x] Added status indicator badge for live sync verification.
- [x] **Configure Data Source**:
  - [x] Linked the SharePoint guest URL for `2026 Fralin Acknowledgements.xlsx`.
  - [x] Targeted the `Donation Screen` tab (31 donors).
  - [x] Generated `donors.csv` in the repository for offline / instant loading.
- [x] **Test End-to-End**:
  - [x] Verified parser and rendering of the 31 donors from `Donation Screen`.


