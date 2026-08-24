# Fralin Development - Live Donor Roll

A responsive, dynamically scrolling donor recognition display for the Fralin Museum of Art. Supports live data from Google Sheets, SharePoint / Excel Online (`.xlsx`), or local CSV files.

---

## 🎨 How to Embed in Canva (Presentations, Digital Signage & Websites)

Canva has a built-in **Embeds** tool that lets you paste any live web link directly into your design.

### 1. The Best Links for Canva:

- **Option A: White Background with `#127CC2` Text in Poppins**
  ```text
  https://fralin-development.github.io/?theme=light
  ```

- **Option B: Transparent Background with `#127CC2` Text (Best for placing over Canva graphics/backgrounds!)**
  ```text
  https://fralin-development.github.io/?theme=light&bg=transparent
  ```

- **Option C: Transparent Background + No Title (Use if you already added a title in Canva)**
  ```text
  https://fralin-development.github.io/?theme=light&bg=transparent&header=false
  ```

- **Option D: Dark Gala Museum Theme**
  ```text
  https://fralin-development.github.io/?embed=true
  ```

### 2. Steps to Add into Canva:
1. Open your design or presentation in [Canva](https://www.canva.com).
2. On the left-hand toolbar, click **Apps** → search for **Embeds** (or click the **Embeds** icon).
3. Paste one of the URLs above into the box and click **Add to design**.
4. Resize and drag the scroll box anywhere on your Canva slide/page!

---

## 🎓 HTML / iframe Embed Codes (For Webpages & LMS)

If you are embedding using raw HTML / `<iframe>`:

```html
<!-- Clean Light Theme (Poppins + #127CC2) -->
<iframe 
    src="https://fralin-development.github.io/?embed=true&theme=light" 
    width="100%" 
    height="600" 
    style="border: 1px solid #e2e8f0; border-radius: 12px; max-width: 900px; display: block; margin: 0 auto;" 
    allowfullscreen>
</iframe>
```

---

## ⚙️ URL Customization Parameters

| Parameter | Example | Description |
| :--- | :--- | :--- |
| `theme=light` | `?theme=light` | White background with `#127CC2` Poppins text |
| `bg=transparent` | `?bg=transparent` | Removes background color so Canva slide background shows through |
| `header=false` | `?header=false` | Hides the top "Special Thanks to Our Donors" header |
| `speed=35` | `?speed=35` | Sets scroll speed in pixels/second (default: 45) |
| `title=...` | `?title=Honor%20Roll` | Dynamically overrides the main header title |
| `subtitle=...` | `?subtitle=2026` | Adds or overrides a subtitle |


