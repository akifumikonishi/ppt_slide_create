/**
 * Crypto Price Viewer - Main JavaScript
 */

// Utility functions
const CryptoViewer = {
    // Format number with thousand separators
    formatNumber: function(num) {
        return new Intl.NumberFormat('ja-JP').format(Math.round(num));
    },

    // Format percentage
    formatPercent: function(num) {
        const sign = num >= 0 ? '+' : '';
        return sign + num.toFixed(2) + '%';
    },

    // Format currency
    formatCurrency: function(num, symbol = '¥') {
        return symbol + this.formatNumber(num);
    },

    // Add flash animation to element
    flashElement: function(element) {
        element.classList.add('price-updated');
        setTimeout(() => {
            element.classList.remove('price-updated');
        }, 1000);
    },

    // Show toast notification
    showToast: function(message, type = 'info') {
        const toastContainer = document.getElementById('toast-container');
        if (!toastContainer) {
            const container = document.createElement('div');
            container.id = 'toast-container';
            container.className = 'position-fixed bottom-0 end-0 p-3';
            container.style.zIndex = '1050';
            document.body.appendChild(container);
        }

        const toast = document.createElement('div');
        toast.className = `toast show bg-${type === 'error' ? 'danger' : type === 'success' ? 'success' : 'info'} text-white`;
        toast.innerHTML = `
            <div class="toast-body">
                ${message}
            </div>
        `;

        document.getElementById('toast-container').appendChild(toast);

        setTimeout(() => {
            toast.remove();
        }, 3000);
    },

    // Debounce function
    debounce: function(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }
};

// Initialize tooltips if Bootstrap is loaded
document.addEventListener('DOMContentLoaded', function() {
    // Initialize Bootstrap tooltips
    const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function(tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });

    console.log('Crypto Price Viewer initialized');
});
