# Adding a New Demo

To add a new demo to the Colibri Alpha gallery:

## Step 1: Create Demo Directory
Create a new folder in the `demos/` directory with a descriptive name:
```
demos/your-demo-name/
```

## Step 2: Add Your Demo Files
Add your demo files to the directory. At minimum, you need:
- `index.html` - The main demo page

You can also include:
- CSS files for styling
- JavaScript files for functionality
- Images or other assets
- Additional HTML pages

## Step 3: Update demos.json
Add your demo information to the `demos.json` file in the root directory:

```json
{
  "title": "Your Demo Title",
  "description": "A brief description of what your demo showcases",
  "url": "./demos/your-demo-name/index.html",
  "icon": "🎯",
  "status": "Alpha",
  "tags": ["tag1", "tag2"],
  "dateAdded": "2024-01-01",
  "version": "0.1.0"
}
```

### Available Status Options:
- `Alpha` - Early development version
- `Beta` - More stable, ready for broader testing
- `Live` - Production-ready feature
- `Stable` - Fully tested and stable

## Step 4: Commit and Push
Commit your changes and push to GitHub. The demo will automatically appear in the gallery!

## Demo Template
Use the sample demo as a template for consistent styling and navigation. Key elements to include:

1. **Back link**: `<a href="../../index.html">← Back to Gallery</a>`
2. **Consistent styling**: Use similar color schemes and fonts
3. **Responsive design**: Ensure it works on mobile devices
4. **Clear title and description**: Help users understand what they're viewing

## Example Demo Structure
```
demos/
  your-demo-name/
    index.html          # Main demo page
    style.css          # Demo-specific styles (optional)
    script.js          # Demo-specific JavaScript (optional)
    assets/            # Images, fonts, etc. (optional)
      image1.png
      icon.svg
```

## Tips
- Keep demos simple and focused on one feature or concept
- Use meaningful names for your demo directories
- Test your demo before committing
- Include clear instructions if the demo requires user interaction
- Consider adding a brief explanation of what the demo demonstrates