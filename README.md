# colibri-alpha  
Alpha versions of browser and software elements for testing and iterative design.

## 🚀 Live Demo Gallery
Visit the live demo gallery at:  
https://esteve32.github.io/colibri-alpha/

## 📝 About
This repository hosts alpha versions and experimental demos for Colibri projects.  
The main site provides a branded gallery interface where visitors can browse and launch in-progress prototypes directly in the browser.

## 🎯 Quick Start
- Open the live gallery to explore current demos.
- Click any demo card to launch the prototype.
- Use in-app navigation to return to the gallery and continue testing.

## ➕ Adding New Demos
To add a new demo to the gallery:

1. Create a demo folder under `demos/` (for example: `demos/your-demo-name/`).
2. Add or update your demo metadata in `demos.json`.
3. Commit and push your changes — the gallery reads from the JSON configuration and surfaces new entries automatically.

## 🏗️ Repository Structure
```text
├── .github/            # Repo-level GitHub configuration
├── README.md           # Project documentation
├── demos.json          # Demo catalog/configuration data
├── demos/              # Individual demo projects and assets
├── google-sheets/      # Google Sheets-related integration/assets
├── index.html          # Main gallery page and app shell
├── scripts.js          # Core gallery behavior and demo loading logic
└── styles.css          # Gallery UI styling, layout, and visual theme
```

## 🎨 Features
- **Central demo gallery:** Single entry point for discovering alpha builds.
- **Config-driven listing:** `demos.json` controls what appears in the UI.
- **Vanilla front-end stack:** Lightweight HTML/CSS/JavaScript implementation.
- **GitHub Pages deployment:** Public and easy-to-share preview environment.
- **Extensible demo architecture:** New demos can be added without rewriting the gallery framework.

## 🔧 Technical Details
- **Hosting:** GitHub Pages  
- **Frontend:** HTML, CSS, JavaScript  
- **Demo management:** JSON-based configuration (`demos.json`)  
- **Project organization:** Root gallery app + modular demo directories (`demos/`) + auxiliary integration area (`google-sheets/`)  
