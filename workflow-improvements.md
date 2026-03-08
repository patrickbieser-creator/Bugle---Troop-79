# Bugle Workflow Improvement Ideas

Discussed: 2026-02-28

---

## 1. Image Upload Script (Highest Priority)

**Problem:** Uploading images to Bunny CDN is multi-step — upload via UI, navigate to file, copy URL, switch windows, type markdown syntax.

**Solution:** A PowerShell script that:
- Accepts an image file as input (drag-and-drop or argument)
- Uploads it to Bunny CDN via their Storage API
- Automatically copies the finished markdown syntax to the clipboard

```
![alt text](https://Troop79.b-cdn.net/filename.png)
```

**Usage idea:**
```powershell
.\Upload-Image.ps1 -ImagePath "C:\Photos\WinterCamp.jpg" -AltText "Winter Camp 2026"
# → copies ![Winter Camp 2026](https://Troop79.b-cdn.net/WinterCamp.jpg) to clipboard
```

**Requires:** Bunny CDN Storage API key and storage zone name.

---

## 2. New Issue Scaffold Script

**Problem:** Each week requires manually copying last week's `.md`, clearing content, updating the date in `Go.ps1`, and updating the hero image URL.

**Solution:** A `New-Bugle.ps1` script that:
- Archives `Bugle.md` → `Bugle YYYY-MM-DD.md` (using last week's date)
- Resets `Bugle.md` from a clean weekly skeleton (date pre-filled to this week)
- Updates the `-BugleDate` in `Go.ps1` automatically
- Optionally prompts for the hero image URL

**Usage idea:**
```powershell
.\New-Bugle.ps1 -Date "March 1, 2026" -HeroImage "https://Troop79.b-cdn.net/hero.png"
```

---

## 3. Meeting Section Fill-in Snippet

**Problem:** The meeting section always has the same structure but you have to remember the format each week.

**Solution:** Keep a snippet template in the repo:

```markdown
## Troop Meeting This Sunday
**Sunday**, [DATE], **4:00–5:30**
**Uniform:** [CLASS A or CLASS B]
![Class B Uniform](https://Troop79.b-cdn.net/Class-B-Uniform.png){.scout-img-40}
**Location:** [LOCATION]
**Snack:** [SNACK PATROL]
**Flag Ceremony:** [PATROL]
**Cleanup:** [PATROL]
[Weekly Duty Roster]([ROSTER LINK])

### Meeting Plan

#### 1. [ACTIVITY]
[Description]

#### 2. [ACTIVITY]
[Description]
```

Copy from the snippet rather than reconstructing from memory each week.

---

## 4. Static Include Files for Boilerplate Sections

**Problem:** "Check Your Clipboard" (~30 scout links) and the merit badge tracking list (~30 badge links) are largely unchanged week to week but take up most of `Bugle.md`.

**Solution:** Store them as separate files:
- `includes/clipboards.md`
- `includes/merit-badges.md`

Have `Build-Bugle.ps1` stitch them in at build time using a placeholder like `{{INCLUDE:clipboards}}`. Only edit these files when the content actually changes.

**Benefit:** `Bugle.md` becomes short and focused on just the weekly content.

---

## 5. Calendar as Simple Data Format

**Problem:** Editing the calendar requires manually writing raw HTML `<tr>` and `<td>` tags, which is tedious and error-prone.

**Solution:** Write the calendar in a simple pipe-delimited `.txt` or `.csv` file:

```
Feb | 22      | Troop Meeting      | Robotics Part III / Trail to Eagle
Mar | 1       | PLC Meeting        | SPLs and Patrol Leaders 3:00
Mar | 1       | Troop Meeting      | Trail to Eagle / Open Advancement
Mar | 6 (Fri) | Committee Meeting  | Unity Lutheran Church 6:00pm
```

Have `Build-Bugle.ps1` (or a separate step) convert this to the styled `<table>` HTML with the correct CSS classes (`month-cell`, `event-cell`).

**Month header rows** can be detected automatically when the month changes.

---

## Implementation Order Recommendation

| Priority | Item | Effort | Time Saved |
|---|---|---|---|
| 1 | Image Upload Script | Medium | High |
| 2 | New Issue Scaffold Script | Low | Medium |
| 3 | Meeting Snippet | None (just save a file) | Low-Medium |
| 4 | Static Includes | Medium | Medium |
| 5 | Calendar Data Format | High | Medium |
