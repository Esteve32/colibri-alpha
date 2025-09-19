# Price Calculator

A web-based price calculator application for calculating product prices with discounts and taxes.

## Features

- Product selection from a predefined list
- Quantity input
- Discount percentage calculation
- Tax rate calculation
- Real-time price updates
- Responsive design
- Clean, modern interface

## Files

- `index.html` - Main HTML structure
- `styles.css` - CSS styling and responsive design
- `script.js` - JavaScript functionality and calculations
- `products.json` - Product data file (simulates downloaded data)

## How to Use

1. Open `index.html` in a web browser
2. Select a product from the dropdown
3. Enter the desired quantity
4. Adjust discount percentage if applicable
5. Set the tax rate (default 10%)
6. The price calculation updates automatically
7. Click "Calculate Price" button to refresh calculations

## Price Calculation Logic

1. **Base Price**: Individual product price
2. **Subtotal**: Base price × Quantity
3. **Discount**: Subtotal × (Discount % ÷ 100)
4. **After Discount**: Subtotal - Discount amount
5. **Tax**: After discount amount × (Tax % ÷ 100)
6. **Total Price**: After discount amount + Tax amount

## Product Data

The application includes 10 sample products across different categories:
- Electronics (Laptop, Monitor, Smartphone, Tablet, Webcam)
- Accessories (Mouse, Keyboard, USB-C Hub)
- Audio (Headphones)
- Wearables (Smartwatch)

## Technical Details

- Pure HTML, CSS, and JavaScript (no external dependencies)
- Responsive design for mobile and desktop
- Real-time calculations with input validation
- Professional UI with gradient backgrounds and smooth animations