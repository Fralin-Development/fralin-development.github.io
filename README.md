# Fralin Development - Live Donor Roll

A responsive, smoothly scrolling live donor recognition display built for the **Fralin Museum of Art**. 

Default styling features **Poppins typography**, **`#127CC2` vibrant blue text**, and a clean **white background**.

---

## 🖥️ How to Display the Donor Roll

### 1. Fullscreen Kiosk / TV / Event Display (Recommended)
Open the URL directly in any modern browser (Chrome, Safari, Edge, Firefox):
```text
https://fralin-development.github.io/
```
* **Enter Fullscreen**: Press `F11` (Windows) or `Cmd + Shift + F` (Mac).
* **Live Updates**: The page automatically checks for donor list updates every 15 seconds in the background and updates the scrolling roll smoothly without reloading.

---

### 2. Embed in Websites & Intranets (WordPress, SharePoint, Canvas LMS, etc.)
Standard website builders and CMS platforms support direct `<iframe>` embedding:

```html
<iframe 
    src="https://fralin-development.github.io/?embed=true" 
    width="100%" 
    height="650" 
    style="border: none; border-radius: 12px; max-width: 960px; display: block; margin: 0 auto;" 
    allowfullscreen>
</iframe>
```

---

## ⚙️ URL Customization Parameters

You can customize the appearance and behavior by adding query parameters to the URL:

| Parameter | Example | Description |
| :--- | :--- | :--- |
| `bg=transparent` | `?bg=transparent` | Makes the background transparent to blend with underlying page backgrounds |
| `header=false` | `?header=false` | Hides the top banner title |
| `speed=...` | `?speed=35` | Custom scroll speed in pixels per second (default: `45`) |
| `title=...` | `?title=Honor%20Roll` | Overrides the main header title |
| `subtitle=...` | `?subtitle=2026%20Campaign` | Adds or overrides a subtitle |
| `theme=dark` | `?theme=dark` | Activates the dark museum gala theme (gold text on dark navy) |

**Example Combined URL:**
```text
https://fralin-development.github.io/?bg=transparent&speed=35
```

---

## 🔄 Automated SharePoint Sync

The repository includes an automated GitHub Action workflow (`.github/workflows/sync-donors.yml`):
* **Automatic Schedule**: Runs every 5 minutes in the background, downloads the latest Excel workbook from the SharePoint guest link, and extracts names from the **`Donation Screen`** tab.
* **Smart Commits**: If new donors are detected, it updates `donors.csv` and pushes to `main`. If no changes were made, it exits cleanly without creating empty commits.
* **Manual Instant Sync**: You can also trigger an immediate sync at any time by going to the repository's **Actions** tab → **Auto-Sync Donors from SharePoint** → **Run workflow**.
