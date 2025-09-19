// Colibri Alpha Demo Gallery JavaScript

document.addEventListener('DOMContentLoaded', function() {
    loadDemos();
});

async function loadDemos() {
    try {
        const response = await fetch('demos.json');
        if (!response.ok) {
            console.log('No demos.json found, showing default state');
            return;
        }
        
        const data = await response.json();
        renderDemos(data.demos);
    } catch (error) {
        console.log('Error loading demos:', error);
        // Silently fail - the default "coming soon" card will remain
    }
}

function renderDemos(demos) {
    const container = document.getElementById('demos-container');
    
    if (!demos || demos.length === 0) {
        return; // Keep the default "coming soon" card
    }
    
    // Clear the coming soon card
    container.innerHTML = '';
    
    demos.forEach((demo, index) => {
        const demoCard = createDemoCard(demo, index);
        container.appendChild(demoCard);
    });
    
    // Add the "coming soon" card at the end
    const comingSoonCard = createComingSoonCard();
    container.appendChild(comingSoonCard);
}

function createDemoCard(demo, index) {
    const card = document.createElement('div');
    card.className = 'demo-card';
    card.style.animationDelay = `${index * 0.1}s`;
    
    const statusClass = getStatusClass(demo.status);
    
    card.innerHTML = `
        <div class="demo-icon">${demo.icon || '🔬'}</div>
        <h4>${demo.title}</h4>
        <p>${demo.description}</p>
        <div class="demo-status ${statusClass}">${demo.status || 'Alpha'}</div>
        <a href="${demo.url}" class="demo-link" target="_blank" rel="noopener">
            View Demo
        </a>
    `;
    
    return card;
}

function createComingSoonCard() {
    const card = document.createElement('div');
    card.className = 'demo-card coming-soon';
    
    card.innerHTML = `
        <div class="demo-icon">🚀</div>
        <h4>More Demos Coming Soon</h4>
        <p>New alpha features and prototypes will be added regularly. Check back often for the latest updates!</p>
        <div class="demo-status">Coming Soon</div>
    `;
    
    return card;
}

function getStatusClass(status) {
    if (!status) return '';
    
    const statusLower = status.toLowerCase();
    if (statusLower.includes('live') || statusLower.includes('stable')) {
        return 'live';
    } else if (statusLower.includes('beta')) {
        return 'beta';
    }
    return '';
}

// Add smooth scrolling for anchor links
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            target.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    });
});

// Add some interactive feedback
document.querySelectorAll('.demo-card').forEach(card => {
    card.addEventListener('mouseenter', function() {
        this.style.transform = 'translateY(-5px) scale(1.02)';
    });
    
    card.addEventListener('mouseleave', function() {
        this.style.transform = 'translateY(0) scale(1)';
    });
});