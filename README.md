# Fralin Development - Live Donor Roll

A responsive, dynamically scrolling donor recognition display for the Fralin Museum of Art. Supports live data from Google Sheets, SharePoint / Excel Online (`.xlsx`), or local CSV files.

---

## 🎓 Canvas LMS Embed Snippets

To embed this live donor scroll directly inside a **Canvas Page**, **Announcement**, or **Module**, use the following snippets:

### 1. Default Dark Museum Theme
```html
<iframe 
    src="https://fralin-development.github.io/?embed=true" 
    width="100%" 
    height="600" 
    style="border: none; border-radius: 12px; max-width: 900px; display: block; margin: 0 auto;" 
    allowfullscreen>
</iframe>
```

### 2. Clean Light Theme (Matches Canvas White Background)
```html
<iframe 
    src="https://fralin-development.github.io/?embed=true&theme=light" 
    width="100%" 
    height="600" 
    style="border: 1px solid #e2e8f0; border-radius: 12px; max-width: 900px; display: block; margin: 0 auto;" 
    allowfullscreen>
</iframe>
```

### 3. Transparent Minimal (No Title / Transparent Background)
```html
<iframe 
    src="https://fralin-development.github.io/?embed=true&bg=transparent&header=false" 
    width="100%" 
    height="500" 
    style="border: none; width: 100%;" 
    allowfullscreen>
</iframe>
```

### 4. Adding as a Canvas Module Item
1. In Canvas, navigate to **Modules** → click **`+` (Add Item)**.
2. Select **External URL** from the dropdown.
3. Enter URL: `https://fralin-development.github.io/?embed=true`
4. Enter Page Name: `Fralin Donor Roll`
5. Click **Add Item**.

---

## ⚙️ URL Customization Parameters

| Parameter | Example | Description |
| :--- | :--- | :--- |
| `embed=true` | `?embed=true` | Optimizes viewport and hides status badges for iframe embeds |
| `theme=light` | `?theme=light` | Swatch to light theme (white background, dark typography) |
| `bg=transparent` | `?bg=transparent` | Removes background color to blend with container |
| `header=false` | `?header=false` | Hides the top "Special Thanks" header |
| `speed=35` | `?speed=35` | Sets scroll speed in pixels/second (default: 45) |
| `title=...` | `?title=Honor%20Roll` | Dynamically overrides the main header title |
| `subtitle=...` | `?subtitle=2026` | Adds or overrides a subtitle |

