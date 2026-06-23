# Filashop

**Pellet-to-Filament Recipe Studio** | 颗粒制线配方工作站

Filashop is a professional recipe management tool for filament extrusion production lines — from raw pellets to finished spools. Think of it as **Photoshop for filament recipes**: mix materials, dial in colors, fine-tune extrusion parameters, and preview the result in real-time 3D.

---

## What It Does

### Material Selection
- **Active materials**: PLA, Transparent PLA, PETG
- **Blend mode**: Mix two pellet types with adjustable ratio (e.g. 70% PLA + 30% PETG)
- Frozen materials (ABS, TPU, Nylon, ASA, PC, HIPS, PP, POM) — parameters pre-configured, coming soon

### Coloring
- **Color powder** (色粉) dyeing method
- **Single color** or **mix mode** (2–3 colors with ratio control)
- Subtractive color mixing prediction — see the blended result before extrusion
- Recent colors history

### Extrusion Parameters
Three core production-line parameters:
| Parameter | Description |
|-----------|-------------|
| Barrel Temperature (料桶温度) | Melting zone heat control |
| Screw Speed (螺杆速度) | Extrusion feed rate (RPM) |
| Traction Speed (牵引速度) | Pulling speed for diameter control (mm/s) |

Smart range indicators warn when parameters fall outside recommended values.

### 3D Spool Preview
- Real-time **Three.js** 3D spool model with realistic winding texture
- Drag to rotate, scroll to zoom
- Live color and material updates — see your recipe before you make it
- Transparent material rendering for PLA-T

### Recipe Library
- **Built-in recipes** included (PLA White Standard, PLA+Carbon Fiber, PLA+Coffee Grounds, etc.)
- **Save / Load / Import / Export** recipes as JSON
- Manage your recipe collection like Photoshop presets
- Copy recipe summary to clipboard for sharing

### Keyboard Shortcuts
| Key | Action |
|-----|--------|
| `S` | Save recipe |
| `E` | Export |
| `R` | Reset |
| `Esc` | Close modal |

### Bilingual Support
- 中文 / English toggle in the top bar
- Language preference saved automatically

---

## Try It Online

Visit the live demo:

**https://xahdiyar37-hub.github.io/filament-recipe-studio/**

## Run Locally

```bash
# Clone the repo
git clone https://github.com/xahdiyar37-hub/filament-recipe-studio.git
cd filament-recipe-studio

# Serve with any static server
python3 -m http.server 8000
# Then open http://localhost:8000
```

No build step required — it's a single `index.html` file.

---

## Join the Recipe Community

**Bought a filament extrusion line? We want you!**

Filashop is built for makers and small-scale filament producers. If you own a pellet-to-filament extrusion machine, you can help grow the recipe library:

### How to Contribute Recipes
1. **Use Filashop** to record your production parameters — material blend, color ratios, barrel temperature, screw speed, traction speed
2. **Export your recipe** as JSON using the in-app export button
3. **Share it back** — submit a Pull Request with your recipe JSON, or open an Issue with the recipe details

### What We're Looking For
- Proven recipes you've actually extruded successfully
- Material blend experiments (PLA+PETG, PLA+TPU, etc.)
- Specialty filaments (silk, glow, wood-fill, carbon fiber reinforced)
- Color formulations (especially hard-to-achieve shades)
- Parameter tuning notes for different pellet suppliers

### Recipe Format
Each recipe includes:
- Base material + optional blend material and ratio
- Coloring method, color hex, and dosage percentage
- Additive selections and percentages
- Three extrusion parameters (barrel temp, screw speed, traction speed)

### Guidelines
- Include the **pellet supplier/brand** if possible — material properties vary
- Note the ** filament diameter** you achieved (1.75mm / 2.85mm / 3.00mm)
- Add **ambient conditions** (room temp, humidity) if relevant — they affect extrusion
- Mark experimental recipes clearly

---

## Tech Stack

- **HTML5 + CSS3 + Vanilla JS** — single-file app, no dependencies
- **Three.js** — WebGL 3D rendering (OrbitControls, PBR materials, soft shadows)
- **Canvas API** — dynamic texture generation for spool winding
- **localStorage** — recipe persistence

## License

MIT — free to use, modify, and distribute.

---

## Roadmap

- [ ] Unfreeze additive system (carbon fiber, coffee grounds, fluorescent, metal powder, wood fill, glow powder, carbon black, glass fiber)
- [ ] Unfreeze additional materials (ABS, TPU, Nylon, ASA, PC, HIPS, PP, POM)
- [ ] Recipe versioning and fork history
- [ ] Cloud recipe sharing platform
- [ ] Production batch logging and quality tracking
- [ ] Diameter tolerance calculator

---

*Filashop — from pellets to perfection.*
