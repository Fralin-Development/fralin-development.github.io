# 🏛️ PowerPoint Native Donor Roll Setup

This guide explains how to sync live donors from GitHub Pages directly into PowerPoint as native vector text with a smooth, movie-style rolling "Credits" animation—**without needing any webview add-ins, browser controls, or network share setups**.

---

## 📁 Files
* **[`scripts/PowerPoint_DonorRoll_Sync.bas`](scripts/PowerPoint_DonorRoll_Sync.bas)**: VBA module that handles live downloading, CSV parsing, slide formatting, and credits animation.

---

## 🚀 How to Import and Run in PowerPoint

### 1. Open the VBA Editor
* **Windows**: Press `Alt + F11`
* **Mac**: Press `Option + F11` (or click **Tools** > **Macro** > **Visual Basic Editor**)

### 2. Import the Macro Module
* In the VBA Editor menu, go to **File** > **Import File...**
* Browse to [`scripts/PowerPoint_DonorRoll_Sync.bas`](scripts/PowerPoint_DonorRoll_Sync.bas) and select **Open**.
* *(Alternatively, click **Insert** > **Module** and copy/paste the code).*

### 3. Run the Macro
* Close the VBA Editor to return to PowerPoint.
* Press `Alt + F8` (or go to **View** > **Macros** / **Developer** > **Macros**).
* Choose one of the two macros:
  * **`SyncDonorsAndCreateRollingList`**: Standard light theme (White background, Fralin Blue `#127CC2` typography).
  * **`SyncDonorsDarkMode`**: Gala dark theme (Dark background, Gold `#E5C158` / White typography).
* Click **Run**.

### 4. Play the Slide Show
* Press **`F5`** (or **`Shift + F5`**) to start the presentation.
* The donor list will automatically roll upward smoothly as native vector text!

---

## ⚙️ Customization Options
You can adjust the constants at the top of the VBA module in `scripts/PowerPoint_DonorRoll_Sync.bas`:

```vba
Public Const CSV_URL As String = "https://fralin-development.github.io/donors.csv"
Public Const HEADER_TITLE As String = "THANK YOU TO OUR GENEROUS DONORS"
Public Const SUBHEADER_TITLE As String = "Who Make Art Together Possible"
Public Const FONT_FAMILY As String = "Poppins"
Public Const SCROLL_SPEED_POINTS_PER_SEC As Single = 60  ' Scroll speed in points/sec (60 = smooth & readable)
Public Const ANIM_REPEAT_COUNT As Long = 1000            ' Infinite / continuous looping
Public Const TARGET_SLIDE_INDEX As Long = 1           ' Slide number to update
```
