// Product data - simulating data from a file
const products = [
    { id: 1, name: "Laptop Computer", price: 899.99 },
    { id: 2, name: "Wireless Mouse", price: 29.99 },
    { id: 3, name: "Mechanical Keyboard", price: 129.99 },
    { id: 4, name: "Monitor 24\"", price: 199.99 },
    { id: 5, name: "USB-C Hub", price: 49.99 },
    { id: 6, name: "Webcam HD", price: 79.99 },
    { id: 7, name: "Headphones", price: 159.99 },
    { id: 8, name: "Smartphone", price: 699.99 },
    { id: 9, name: "Tablet", price: 329.99 },
    { id: 10, name: "Smartwatch", price: 249.99 }
];

// DOM elements
const productSelect = document.getElementById('product-select');
const quantityInput = document.getElementById('quantity');
const discountInput = document.getElementById('discount');
const taxInput = document.getElementById('tax');
const calculateBtn = document.getElementById('calculate-btn');

// Result elements
const basePriceSpan = document.getElementById('base-price');
const subtotalSpan = document.getElementById('subtotal');
const discountAmountSpan = document.getElementById('discount-amount');
const afterDiscountSpan = document.getElementById('after-discount');
const taxAmountSpan = document.getElementById('tax-amount');
const totalPriceSpan = document.getElementById('total-price');

// Initialize the application
function init() {
    populateProductSelect();
    attachEventListeners();
    calculatePrice(); // Initial calculation
}

// Populate the product select dropdown
function populateProductSelect() {
    products.forEach(product => {
        const option = document.createElement('option');
        option.value = product.id;
        option.textContent = `${product.name} - $${product.price.toFixed(2)}`;
        productSelect.appendChild(option);
    });
}

// Attach event listeners
function attachEventListeners() {
    calculateBtn.addEventListener('click', calculatePrice);
    productSelect.addEventListener('change', calculatePrice);
    quantityInput.addEventListener('input', calculatePrice);
    discountInput.addEventListener('input', calculatePrice);
    taxInput.addEventListener('input', calculatePrice);
}

// Main calculation function
function calculatePrice() {
    const selectedProductId = parseInt(productSelect.value);
    const quantity = parseInt(quantityInput.value) || 0;
    const discountPercent = parseFloat(discountInput.value) || 0;
    const taxPercent = parseFloat(taxInput.value) || 0;

    // Find selected product
    const selectedProduct = products.find(p => p.id === selectedProductId);
    
    if (!selectedProduct || quantity <= 0) {
        resetResults();
        return;
    }

    // Calculate prices
    const basePrice = selectedProduct.price;
    const subtotal = basePrice * quantity;
    const discountAmount = subtotal * (discountPercent / 100);
    const afterDiscount = subtotal - discountAmount;
    const taxAmount = afterDiscount * (taxPercent / 100);
    const totalPrice = afterDiscount + taxAmount;

    // Update display
    updateResults({
        basePrice,
        subtotal,
        discountAmount,
        afterDiscount,
        taxAmount,
        totalPrice
    });
}

// Update result display
function updateResults(results) {
    basePriceSpan.textContent = formatCurrency(results.basePrice);
    subtotalSpan.textContent = formatCurrency(results.subtotal);
    discountAmountSpan.textContent = `-${formatCurrency(results.discountAmount)}`;
    afterDiscountSpan.textContent = formatCurrency(results.afterDiscount);
    taxAmountSpan.textContent = formatCurrency(results.taxAmount);
    totalPriceSpan.textContent = formatCurrency(results.totalPrice);
}

// Reset results to zero
function resetResults() {
    const zeroValue = formatCurrency(0);
    basePriceSpan.textContent = zeroValue;
    subtotalSpan.textContent = zeroValue;
    discountAmountSpan.textContent = `-${zeroValue}`;
    afterDiscountSpan.textContent = zeroValue;
    taxAmountSpan.textContent = zeroValue;
    totalPriceSpan.textContent = zeroValue;
}

// Format number as currency
function formatCurrency(amount) {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD'
    }).format(amount);
}

// Start the application when DOM is loaded
document.addEventListener('DOMContentLoaded', init);