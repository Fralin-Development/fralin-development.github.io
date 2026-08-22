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

## 2. To-Do List (When You Return)

- [ ] **Commit Pending Files**: Commit `_config.yml` to `dev_steph_soph` and push.
- [ ] **Provide Google Sheet Link / Tab**: Share the public Google Sheet link (or tab URL with `#gid=...`).
- [ ] **Update `script.js`**:
  - Add auto-parser for Google Sheets share/tab URLs.
  - Implement direct CSV fetch + parsing (no API key needed).
  - Handle column mappings (e.g., Donor Name, Amount, Custom Message).
  - Configure auto-refresh interval for live donor updates.
- [ ] **Customize Scrolling & Visuals (`style.css`)**:
  - Adjust animation speed, font sizes, colors, and layout according to display requirements (e.g., projector, kiosk, or web embed).
- [ ] **Test End-to-End**:
  - Add/modify sample rows in the Google Sheet and verify that the scroll updates live without page refreshes.
