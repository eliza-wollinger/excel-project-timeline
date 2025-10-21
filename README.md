# Excel Project Timeline

A simple yet powerful **Office Script** that transforms your Excel sheet into a **visual project timeline**, just like MS Project, but way lighter and built right into Excel.

Perfect for devs, analysts, or project managers who love automation and clean data visuals.

---

## What it does

- 📥 Reads your project data from the **“Atividades”** worksheet  
- 🧾 Generates a new **“Timeline”** worksheet automatically  
- 📅 Displays up to **60 days** of project activities  
- 🎨 Assigns **unique colors** per project  
- 🕓 Highlights **today’s date** for quick reference  
- 🧩 Shows tasks as a **Gantt-style bar chart** made of “■” blocks  
- 🔄 Clears old timelines automatically — no manual cleanup needed  

---

## Input format

Your source sheet (default: `Atividades`) should look like this:

| Project | Task | Status | Start | Finish |
|----------|------|--------|--------|---------|
| Project A | Planning | In Progress | 10/01/2025 | 10/10/2025 |
| Project A | Execution | Not Started | 10/11/2025 | 10/30/2025 |
| Project B | Testing | Completed | 10/05/2025 | 10/12/2025 |

💡 Notes  
- Column names are **case-insensitive**  
- Dates can be **Excel numbers** or **strings (e.g., 10/01/2025)**  

---

## Output: “Timeline” worksheet

When you run the script, it creates a shiny new **“Timeline”** tab like this:

| PROJECT | TASK | STATUS | 10/01 | 10/02 | 10/03 | ... |
|----------|------|--------|-------|-------|-------|-----|
| Project A | Planning | in progress | ■ | ■ | ■ | ... |

✅ Each “■” = one active day  
🎨 Colors by project  
🌟 Today’s date is highlighted in **yellow**

---

## Configuration

| Variable | Default | Description |
|-----------|----------|-------------|
| `mainWorksheetName` | `"Atividades"` | Name of the source worksheet |
| `outputWorksheetName` | `"Timeline"` | Name of the output worksheet |
| `MAX_DAYS` | `60` | How many days to display on the timeline |

---

## Color scheme

- **Project colors:** vibrant tones (coral, violet, turquoise, etc.)  
- **Status colors:**
  - `Completed` → Light gray  
  - `In Progress` → Medium gray  
  - `Not Started` → Dark gray  

---

## How to use

1. Open **Excel Online** (or desktop Excel with Office Scripts enabled)  
2. Go to **Automate → New Script**  
3. Paste the code
4. Click **Run**  
5. Your project timeline appears instantly  

---

## Power tip

You can take it further with **Power Automate**:
- Auto-generate a fresh timeline every week  
- Email the result as a PDF  
- Refresh a Power BI dashboard for full visibility  

---

## License

MIT License — free to use, tweak, and share.  

---

## Tags

`excel` • `office-scripts` • `automation` • `gantt-chart` • `project-tracking` • `timeline` • `power-automate` • `data-visualization`
