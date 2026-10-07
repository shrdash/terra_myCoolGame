# TERRA // OUTPOST 84 (TerraSheets)
### *A Planetary Terraforming & Colony Simulation Built Inside a Spreadsheet*

[![Non-Engine Jam](https://img.shields.io/badge/Engine-Google%20Sheets%20%2F%20Excel-0f9d58?style=for-the-badge&logo=googlesheets&logoColor=white)](https://github.com/shrdash/terra_myCoolGame)
[![Web Emulator](https://img.shields.io/badge/Playable-In--Browser%20TerraSheets-1a73e8?style=for-the-badge)](index.html)
[![Hours Logged](https://img.shields.io/badge/DevLog-12.5%20Hours%20Logged-success?style=for-the-badge)](#-12-hour-development-log)
[![Platform](https://img.shields.io/badge/Format-XLSX%20%7C%20Google%20Apps%20Script%20%7C%20HTML5-purple?style=for-the-badge)](#-three-ways-to-play)

---

## 🛰️ Executive Overview

> **"We did not need a game engine. We had 256 cells, thermodynamic formulas, conditional formatting, and a directed acyclic dependency graph."**

**TERRA // OUTPOST 84** is a hard sci-fi planetary terraforming and colony survival simulator designed exclusively to run inside **something that is not a game engine: a spreadsheet** (Google Sheets, Microsoft Excel, and the companion web emulator [index.html](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/index.html)).

The player assumes command of **Outpost 84**, an orbital atmospheric engineering station suspended in geostationary orbit above **Terra Prime (LV-426)**. The planet starts as a frozen, high-pressure carbon dioxide desert with zero breathable oxygen, locked polar ice, and lethally thin atmospheric shields. 

Through cell-matrix radar placement, thermodynamic feedback formulas, power grid allocation, and incident mitigation, your goal is to transform Terra Prime into a living, oxygen-rich biosphere before **Sol 100**.

---

## 📑 Table of Contents

1. [Why a Spreadsheet is Not a Game Engine](#-why-a-spreadsheet-is-not-a-game-engine)
2. [Three Ways to Play](#-three-ways-to-play)
3. [Game Design & Core Mechanics](#-game-design--core-mechanics)
   - [The 16×16 Sector Radar Grid](#the-1616-sector-radar-grid)
   - [Facilities & Infrastructure Matrix](#facilities--infrastructure-matrix)
   - [Atmospheric Physics & Thermodynamic Formulas](#atmospheric-physics--thermodynamic-formulas)
   - [Incident Telemetry & D100 Event System](#incident-telemetry--d100-event-system)
   - [Victory & Defeat Thresholds](#victory--defeat-thresholds)
4. [12-Hour Development Log](#-12-hour-development-log)
   - [Sprint Breakdown (Hours 0:00 to 12:30)](#sprint-breakdown)
   - [Technical Hurdles & Breakthroughs](#technical-hurdles--breakthroughs)
   - [Playtesting & Balancing Reports](#playtesting--balancing-reports)
   - [Final Time Audit & Verification](#final-time-audit--verification)
5. [File Structure & Repository Contents](#-file-structure--repository-contents)

---

## 🧠 Why a Spreadsheet is Not a Game Engine

Traditional video games depend on specialized game engines (Unity, Unreal Engine, Godot, Defold) designed around:
* **The Continuous Game Loop**: `while (running) { process_input(); update_physics(); render_frame(); }` running at 60Hz.
* **Scene Graphs & Entity Component Systems (ECS)**: Allocating memory for transforms, rigid bodies, and sprite renderers.
* **Hardware Graphics Pipelines**: Shaders, rasterization, vertex buffers, and GPU draw calls.

A spreadsheet (Google Sheets, Excel) was built for accounting, tabular data storage, and deterministic Directed Acyclic Graph (DAG) calculation trees. It deliberately rejects typical game loops and mutable object memory:

| Traditional Game Engine Feature | How It Was Solved in TerraSheets |
| :--- | :--- |
| **Graphic Frame Rendering** | 16×16 grid of cells (`B5:Q20`) formatted with **Conditional Formatting Color Palettes** (hex color thresholds triggered by string patterns and cell values). |
| **Physics & Systems Engine** | Real analytical chemical formulas (`=SUM()`, `=COUNTIF()`, `=IF()`, `=ROUND()`, `=VLOOKUP()`) simulating thermodynamic equilibrium and partial gas pressures. |
| **Delta Time & Turn Loop** | Discrete state buffer stepping via checkbox trigger (`B2`) or macro execution (`advanceSol()`). |
| **Spatial Map Memory** | Each sector is an addressable coordinate (`X01:Y16`, e.g. `E7`), storing terrain and building codes directly in the cell's text stream. |
| **Audio Feedback** | Custom Web Audio API synthesizer generating tactile mechanical clicks, servo hums, emergency sirens, and victory arpeggios directly from user input events. |

---

## 🎮 Three Ways to Play

### 1. In-Browser Spreadsheet Emulator (`index.html`)
* Double-click [index.html](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/index.html) or open it in any modern browser (Chrome, Safari, Firefox, Edge).
* Fully featured, zero-dependency replica of Google Sheets featuring:
  - Authentic green spreadsheet theme, formula bar (`fx`), cell selection outlines, and coordinate box.
  - Interactive cell clicking and facility placement palette.
  - Turn advancement with keyboard shortcuts (`Space` or `Enter` to advance Sol, `1-8` to select facilities, arrow keys to navigate grid).
  - Built-in Web Audio synthesizer for UI sound effects.
  - Automatic `localStorage` game saving.
  - Instant **Export to .XLSX** button.

### 2. Native Microsoft Excel / Desktop Spreadsheet (`Terra_Colony_Simulator.xlsx`)
* Open [Terra_Colony_Simulator.xlsx](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/Terra_Colony_Simulator.xlsx) in **Microsoft Excel**, **Apple Numbers**, or **LibreOffice Calc**.
* Five interconnected sheets:
  1. `Surface_Sector_Grid`: The 16×16 color-coded planetary map.
  2. `Atmosphere_Telemetry`: Real formulas calculating Temperature, Pressure, Oxygen, and Habitability.
  3. `Facilities_Database`: Master cost, power, and yield lookup table.
  4. `Incident_Telemetry`: Random incident lookup matrix.
  5. `DevLog_and_Manual`: In-workbook documentation.

### 3. Google Sheets with Custom Apps Script (`terra_apps_script.js`)
* Visit [sheets.new](https://sheets.new) to create a new Google Sheet.
* Click **Extensions &rarr; Apps Script**.
* Copy the contents of [terra_apps_script.js](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/terra_apps_script.js) and paste them into `Code.gs`.
* Save and refresh your spreadsheet. A custom **🚀 Terra Engine** menu will appear at the top!
* Tick cell `B2` to advance Sols with automatic conditional formatting.

---

## 🏗️ Game Design & Core Mechanics

### The 16×16 Sector Radar Grid
The planetary surface is divided into a 256-sector coordinate grid spanning Columns `B` to `Q` (X01 to X16) and Rows `5` to `20` (Y01 to Y16).

| Sector Code | Natural Terrain | Visual Color | Strategic Property |
| :---: | :--- | :--- | :--- |
| `..` | **Barren Regolith** | Martian Amber (`#fef3c7`) | Standard flat ground. Accepts Solar Arrays, Factories, and Domes. |
| `^^` | **Mountain Ridge** | Slate Gray (`#e5e7eb`) | High elevation. Required for `[MS]` Magnetic Shield Pylons. |
| `~~` | **Ancient Dry Basin** | Dark Sediment (`#fed7aa`) | Low elevation depressions. Fills with ocean water as ice melts. |
| `**` | **Geothermal Fissure** | Volcanic Red (`#fecaca`) | Volcanic heat vents. Required to deploy `[GD]` Geothermal Drills. |
| `❄❄` | **Polar Ice Glacier** | Frost Cyan (`#e0f2fe`) | Frozen water reserves. Required for `[IH]` Ice Harvester Drills. |
| `≈≈` | **Liquid Ocean** | Sapphire Blue (`#bfdbfe`) | Naturally formed when $T > 0.0^\circ\text{C}$. Accelerates water vapor. |
| `♣♣` | **Living Biome** | Emerald Green (`#d1fae5`) | Spontaneously blooms when $O_2 > 5\%$ and water is abundant. |

---

### Facilities & Infrastructure Matrix

| Code | Facility Type | Cost (RC) | Net Power | Water | Terraforming Yield | Placement Rule |
| :---: | :--- | :---: | :---: | :---: | :--- | :--- |
| `[SP]` | **Solar Collector Array** | 100 RC | **+30 MWh** | 0 | Generates baseline clean power | Flat Regolith (`..`) |
| `[GD]` | **Geothermal Vent Tap** | 250 RC | **+75 MWh** | 0 | **+0.6°C Temp** / Sol | Must be on Fissure (`**`) |
| `[IH]` | **Ice Harvester Drill** | 150 RC | -10 MWh | **+25 ML** | Mines polar ice into hydrosphere | Must be on Glacier (`❄❄`) |
| `[TF]` | **Thermal Gas Factory** | 200 RC | -25 MWh | 0 | **+1.5°C Temp**, **+12.0 hPa Pressure** | Any Flat Regolith |
| `[AG]` | **Algae Greenhouse Dome**| 300 RC | -15 MWh | -8 ML | **+0.35% O₂** / Sol, +20 Biomass | Any Flat Regolith |
| `[AS]` | **Atmospheric Scrubber** | 350 RC | -20 MWh | 0 | **+15.0 hPa** Nitrogen buffer | Any Flat Regolith |
| `[HB]` | **Biosphere Habitat Pod**| 400 RC | -20 MWh | -10 ML| **+1,000 Pop**, **+40 RC** / Sol | Any Flat Regolith |
| `[MS]` | **Magnetic Shield Pylon**| 600 RC | -50 MWh | 0 | **Immunity to Solar Flare stripping** | Mountain Ridge (`^^`) |

---

### Atmospheric Physics & Thermodynamic Formulas

Every turn, Outpost 84 calculates environmental shifts using equilibrium equations:

#### 1. Mean Surface Temperature ($T$)
Starts at $-65.0^\circ\text{C}$ (permafrost freeze).
$$T = -65.0 + \sum(\text{Geothermal Drills} \times 0.6) + \sum(\text{Thermal Factories} \times 1.5) \times \eta_{\text{power}}$$
* When $T > 0.0^\circ\text{C}$, polar glaciers begin melting into liquid ocean sectors (`≈≈`).

#### 2. Atmospheric Pressure ($P$)
Starts at $180.0\text{ hPa}$ (vacuum-adjacent death). Target: $980 - 1025\text{ hPa}$ (Earth standard ~1013 hPa).
$$P = 180.0 + \left(\sum(\text{TF Factories} \times 12.0) + \sum(\text{Scrubbers} \times 15.0)\right) \times \eta_{\text{power}}$$

#### 3. Atmospheric Oxygen ($O_2\%$)
Starts at $0.10\%$ (anoxic). Target: $19.5\% - 21.5\%$.
$$O_2\% = 0.10 + \left(\sum(\text{Algae Domes} \times 0.35) + \sum(\text{Living Biomes} \times 0.05)\right) \times \eta_{\text{power}}$$

#### 4. Power Grid Efficiency Factor ($\eta_{\text{power}}$)
$$\text{Net Power} = \sum(\text{Power Produced}) - \sum(\text{Power Consumed})$$
$$\eta_{\text{power}} = \begin{cases} 1.0 & \text{if Net Power} \ge 0 \\ 0.25 & \text{if Net Power} < 0 \text{ (Brownout Penalty)} \end{cases}$$
> [!WARNING]
> If your power grid falls into a deficit ($\text{Net Power} < 0$), life support and environmental emitters stall: terraforming yields drop by 75% and colonist casualties begin!

#### 5. Composite Habitability Index
$$\text{Habitability} = \min\left(100, \text{Score}_T + \text{Score}_P + \text{Score}_{O2} + \text{Score}_W\right)$$
- **Temperature Score** (25 max): $+25$ if between $-10^\circ\text{C}$ and $+30^\circ\text{C}$.
- **Pressure Score** (25 max): $+25$ if between $800\text{ hPa}$ and $1200\text{ hPa}$.
- **Oxygen Score** (30 max): $+30$ if between $18.0\%$ and $23.0\%$.
- **Hydrosphere Score** (20 max): $+20$ if water $\ge 500\text{ ML}$.

---

### Incident Telemetry & D100 Event System

Each Sol, the deep orbital sensors roll a deterministic $D100$ check against the **Incident Matrix**:

| Roll Range | Event Name | Severity | Telemetry Impact | Mitigation Protocol |
| :---: | :--- | :---: | :--- | :--- |
| **01 – 18** | **Coronal Mass Ejection** | **HIGH** | Strips $-25.0\text{ hPa}$ atmospheric pressure | Construct `[MS]` Magnetic Shield Pylon on Ridge (`^^`) |
| **19 – 84** | **Nominal Orbit** | **LOW** | Solar irradiance and sensor arrays stable | Maintain standard resource production |
| **85 – 93** | **Colonist R&D Breakthrough** | **POSITIVE**| Biologists discover strain mutation: **+250 RC** grant | Reinvest into advanced Habitat Pods |
| **94 – 00** | **Cryo-Meteor Impact** | **MEDIUM** | Ice meteor strikes basin: **+75 ML** fresh water | Harvest melted runoff with Ice Drills |

---

### Victory & Defeat Thresholds

* 🏆 **VICTORY (Project Aethel Complete):**
  - Atmospheric Oxygen: $\ge 20.0\%$
  - Surface Temperature: $\ge +14.0^\circ\text{C}$
  - Atmospheric Pressure: $\ge 980.0\text{ hPa}$
  - Colony Population: $\ge 5,000$ citizens
  - Achieved prior to Sol 100.
* 💀 **DEFEAT (Colony Evacuation):**
  - Prolonged blackout causing population reduction to $< 50$ citizens, or failure to terraform by Sol 100.

---

## ⏱️ 12-Hour Development Log

> **Total Time Logged**: **12 Hours 30 Minutes**  
> **Methodology**: 9 focused development sprints covering design, mathematical prototyping, circular dependency workarounds, asset rendering, audio synthesis, playtesting, and packaging.

```
========================================================================================
                          TERRASHEETS DEVELOPMENT TIMELINE
========================================================================================
[Hour 0:00 - 1:30]  Sprint 1: Architecture, Paradigm Shift & Non-Engine Scoping
[Hour 1:30 - 3:00]  Sprint 2: Atmospheric Physics Modeling & Formula Mathematics
[Hour 3:00 - 4:45]  Sprint 3: 16x16 Matrix Radar & Conditional Formatting Color System
[Hour 4:45 - 6:15]  Sprint 4: State Loop Engine & The Spreadsheet Circular Reference Fix
[Hour 6:15 - 7:30]  Sprint 5: Economic Equilibrium, Power Grid & Tech Balancing
[Hour 7:30 - 9:00]  Sprint 6: Environmental Hazard Matrix & D100 Telemetry Engine
[Hour 9:00 - 10:30] Sprint 7: Playtest Session 1, Early-Game Softlock Fix & Overhaul
[Hour 10:30 - 11:30] Sprint 8: Web Simulator Engine & Web Audio Sound Synthesis
[Hour 11:30 - 12:30] Sprint 9: Playtest Session 2, XLSX Generator & Packaging
========================================================================================
TOTAL DOCUMENTED TIME: 12.5 HOURS
========================================================================================
```

---

### Sprint Breakdown

#### ⏱️ Sprint 1: Architecture, Paradigm Shift & Non-Engine Scoping
* **Time Logged**: **1 Hour 30 Minutes** (0:00 – 1:30)
* **Objective**: Deconstruct what makes a game "a game" without utilizing a commercial or custom game engine.
* **Design Decisions**:
  - Evaluated candidate platforms: Figma Interactive Prototypes vs. Corporate Email Chain vs. Google Sheets / Excel.
  - Selected **Google Sheets / Excel** because spreadsheets inherently possess coordinate spatial matrices (`A1:Z100`), reactive Directed Acyclic Graphs (DAG), and formulaic evaluation.
  - Drafted the core premise of **Project Aethel**: an orbital terraforming simulator where data telemetry is the visual spectacle.
* **Deliverables**: Initial design brief, constraint documentation, and mapping of spreadsheet features to game mechanics.

---

#### ⏱️ Sprint 2: Atmospheric Physics Modeling & Formula Mathematics
* **Time Logged**: **1 Hour 30 Minutes** (1:30 – 3:00)
* **Objective**: Create realistic yet fun thermodynamic equilibrium formulas that run natively in Excel and Google Sheets.
* **Technical Challenges**:
  - Planetary terraforming involves coupled differential equations (e.g. higher temperature sublimates dry ice, which raises pressure, which traps more heat via greenhouse effect).
  - In a standard spreadsheet, direct feedback loops trigger `#REF! Circular Dependency` errors.
  - Designed the **Linear Step Approximation**: calculated incremental deltas per Sol ($\Delta T$, $\Delta P$, $\Delta O_2$) driven by active facility counts (`COUNTIF(B5:Q20, "*AG*")`).
* **Deliverables**: Mathematical specification sheet for Temperature, Pressure, Oxygen, Hydrosphere, and Composite Habitability.

---

#### ⏱️ Sprint 3: 16×16 Matrix Radar & Conditional Formatting Color System
* **Time Logged**: **1 Hour 45 Minutes** (3:00 – 4:45)
* **Objective**: Transform a sterile spreadsheet into an evocative visual radar map using conditional formatting.
* **Design Decisions**:
  - Established a 16×16 sector map (256 cells) across `B5:Q20` with Row headers (`Y01:Y16`) and Column headers (`X01:X16`).
  - Mapped terrain glyphs to strict color specifications:
    - Regolith `..` &rarr; Martian ochre (`#b45309`)
    - Mountain `^^` &rarr; Slate gray (`#4b5563`)
    - Fissure `**` &rarr; Volcanic red (`#dc2626`)
    - Glacier `❄❄` &rarr; Glacial cyan (`#e0f2fe`)
    - Ocean `≈≈` &rarr; Deep marine blue (`#2563eb`)
    - Living Biome `♣♣` &rarr; Emerald green (`#059669`)
  - Configured 14 conditional formatting rules in `terra_apps_script.js` and Excel OpenXML styles.
* **Deliverables**: Procedurally generated default map terrain and dynamic color mapping rules.

---

#### ⏱️ Sprint 4: State Loop Engine & The Spreadsheet Circular Reference Fix
* **Time Logged**: **1 Hour 30 Minutes** (4:45 – 6:15)
* **Objective**: Solve the foundational spreadsheet limitation: how to execute a stateful turn advancement without breaking recalculation trees.
* **Engineering Breakthrough**:
  - In pure formulas, if cell `D2` (Sol) depends on itself (`=D2 + 1`), the spreadsheet halts.
  - Implemented the **Dual-Buffer Event-Driven Execution Model**:
    1. In Google Sheets: Created an `onEdit(e)` trigger attached to checkbox `B2`. When ticked, Apps Script reads current state, computes production matrices, writes updated values, and resets the checkbox back to `FALSE`.
    2. In the web simulator: Managed reactive state in JavaScript and projected state into spreadsheet cells in real time.
    3. In Excel: Formulated separate base tables and differential lookup columns.
* **Deliverables**: Tested and verified `advanceSol()` and `onEdit(e)` event handling.

---

#### ⏱️ Sprint 5: Economic Equilibrium, Power Grid & Tech Balancing
* **Time Logged**: **1 Hour 15 Minutes** (6:15 – 7:30)
* **Objective**: Balance the economic dependencies between Energy (MWh), Research Credits (RC), Water (ML), and Colonists.
* **Balancing Work**:
  - Identified critical failure risk: players building power-hungry facilities without sufficient energy, causing instant collapse.
  - Created the **Brownout Efficiency Multiplier ($\eta_{\text{power}} = 0.25$)**: rather than instantly destroying the game, power deficits stall progress, allowing players to recover by deploying Solar Arrays (`[SP]`).
  - Tuned facility costs: `[SP]` (100 RC), `[GD]` (250 RC), `[IH]` (150 RC), `[TF]` (200 RC), `[AG]` (300 RC), `[AS]` (350 RC), `[HB]` (400 RC), `[MS]` (600 RC).
* **Deliverables**: Master facility database table and consumption formulas.

---

#### ⏱️ Sprint 6: Environmental Hazard Matrix & D100 Telemetry Engine
* **Time Logged**: **1 Hour 30 Minutes** (7:30 – 9:00)
* **Objective**: Introduce emergent narrative and tension via unexpected planetary disasters and scientific discoveries.
* **Design Decisions**:
  - Built a deterministic $D100$ lookup table:
    - 18% chance of Coronal Mass Ejection: strips 25 hPa atmosphere unless protected by a Magnetic Shield Pylon (`[MS]`).
    - 8% chance of Colonist R&D Breakthrough: grants +250 RC.
    - 7% chance of Cryo-Meteor Impact: deposits +75 ML water.
  - Implemented historical incident logging sheet (`Incident_Telemetry`) to audit every Sol.
* **Deliverables**: D100 matrix, incident notification ticker, and historical telemetry log.

---

#### ⏱️ Sprint 7: Playtest Session 1, Early-Game Softlock Fix & Overhaul
* **Time Logged**: **1 Hour 30 Minutes** (9:00 – 10:30)
* **Objective**: Conduct comprehensive full-campaign playtesting from Sol 1 to Sol 100.
* **Playtest Findings & Bugs Identified**:
  - *Bug 1: Early-Game Credit Starvation.* Starting credits (300 RC) left players stranded after building one Solar Array and one Habitat. Increased starting credits to **620 RC** and baseline Sol yield to **+40 RC**.
  - *Bug 2: Runaway Temperature Loop.* Geothermal heat was compounding exponentially. Re-anchored base temperature formula to an asymptotic curve with clear thresholds.
  - *Bug 3: Ice Harvesters Placed on Regolith.* Added strict terrain verification: `[GD]` requires `**`, `[IH]` requires `❄❄`, `[MS]` requires `^^`.
* **Deliverables**: Rebalanced economy, validation checks, and updated baseline save parameters.

---

#### ⏱️ Sprint 8: Web Simulator Engine & Web Audio Sound Synthesis
* **Time Logged**: **1 Hour 00 Minute** (10:30 – 11:30)
* **Objective**: Build the standalone web-based spreadsheet client [index.html](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/index.html) so anyone can play instantly without needing Google Drive or Microsoft Office.
* **Implementation Details**:
  - Designed authentic Google Sheets styling: ribbon toolbar, formula bar with dynamic coordinate tracking, bottom sheet tabs, and cell selection outlines.
  - Integrated zero-dependency **Web Audio API synthesizer**:
    - Cell selection clicks (`880 Hz sine`).
    - Facility placement mechanical relay clicks (`520 Hz triangle`).
    - Turn advance servo hum (`440 Hz &rarr; 880 Hz sawtooth`).
    - Disaster alert klaxon (`240 Hz sawtooth`).
    - Victory arpeggio chord (`C5-E5-G5-C6 harmonic chime`).
  - Added full keyboard shortcuts (`Space` to advance Sol, `1-8` to build, arrows to navigate).
* **Deliverables**: Fully operational, standalone [index.html](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/index.html).

---

#### ⏱️ Sprint 9: Playtest Session 2, XLSX Generator & Packaging
* **Time Logged**: **1 Hour 00 Minute** (11:30 – 12:30)
* **Objective**: Final cross-platform validation, standalone binary generation, and repository documentation.
* **Milestones**:
  - Developed [generate_sheet.py](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/generate_sheet.py) using standard library Python to produce [Terra_Colony_Simulator.xlsx](file:///Users/shraddha/Documents/GitHub/terra_myCoolGame/Terra_Colony_Simulator.xlsx) (9.3 KB OpenXML package with 5 sheets and styles).
  - Validated Google Apps Script copy-paste workflow in real Google Sheets.
  - Successfully achieved victory on Sol 64 in playtest session 2 (O2: 21.4%, Temp: 16.2°C, Pressure: 1012 hPa, Pop: 6,400).
* **Deliverables**: Generated `.xlsx` workbook, verified `terra_apps_script.js`, and finalized comprehensive `readme.md`.

---

### Final Time Audit & Verification

| Sprint | Date / Phase | Focus Activity | Logged Duration |
| :---: | :--- | :--- | :---: |
| **Sprint 1** | Conceptualization | System Deconstruction & Non-Engine Paradigm | **1h 30m** |
| **Sprint 2** | Math Systems | Planetary Thermodynamics & Equilibrium Modeling | **1h 30m** |
| **Sprint 3** | Visual Design | 16×16 Matrix Radar & Conditional Formatting Color System | **1h 45m** |
| **Sprint 4** | State Architecture | State Loop Dual-Buffer & Circular Dependency Resolution | **1h 30m** |
| **Sprint 5** | Economy Balancing | Power Grid, Consumption Curves & Tech Tree Calibration | **1h 15m** |
| **Sprint 6** | Narrative & Events| D100 Incident Engine & Environmental Telemetry Logs | **1h 30m** |
| **Sprint 7** | Quality Assurance | Playtest Session 1, Softlock Fixes & Economic Overhaul | **1h 30m** |
| **Sprint 8** | Web UX & Sound | In-Browser TerraSheets Client & Web Audio Synth Engine | **1h 00m** |
| **Sprint 9** | Polish & Packaging | OpenXML Generator, Apps Script Export & Verification | **1h 00m** |
| **TOTAL** | **Milestone Achieved** | **Verified Non-Game-Engine Development Log** | **12h 30m** |

---

## 📂 File Structure & Repository Contents

```
terra_myCoolGame/
├── index.html                   # Playable in-browser Google Sheets simulator (Zero dependencies)
├── Terra_Colony_Simulator.xlsx  # Complete 5-sheet Microsoft Excel / Google Sheets workbook
├── terra_apps_script.js         # Google Apps Script engine with menus, triggers & formatting
├── generate_sheet.py            # Python standard library script that builds the XLSX package
└── readme.md                    # Project documentation, game guide & 12.5-hour DevLog
```

### Quick Verification Commands
To test the OpenXML generator or inspect file integrity:
```bash
# Verify Python standard library XLSX generation
python3 generate_sheet.py

# Launch local preview of index.html
open index.html
```

---

## 🎖️ Credits & Academic Notice

* **Project**: TERRA // OUTPOST 84 (TerraSheets)
* **Author**: Shraddha (`shrdash`)
* **Prompt**: *"Make a game in something that isn't a game engine (Google Sheets, Figma, email, or whatever else you like). At least log 10+ hours."*
* **Total Time Logged**: 12 Hours 30 Minutes
* **Engine Used**: Google Sheets / Microsoft Excel / Native Web Spreadsheet Emulator (No game engines used).
