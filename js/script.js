(function () {
    const form = document.getElementById('licenseForm');
    if (!form) return;

    const submitBtn = document.getElementById('submitBtn');

    // Auto-select product from URL query param
    const params = new URLSearchParams(window.location.search);
    const productParam = params.get('product');
    if (productParam) {
        const sel = document.getElementById('product');
        if (sel) sel.value = productParam;
    }

    // Toast notification system
    function showToast(message, type) {
        type = type || 'info';
        var container = document.getElementById('toast-container');
        if (!container) {
            container = document.createElement('div');
            container.id = 'toast-container';
            document.body.appendChild(container);
        }
        var icons = { success: 'bx bx-check-circle', error: 'bx bx-x-circle', info: 'bx bx-info-circle' };
        var toast = document.createElement('div');
        toast.className = 'toast toast-' + type;
        toast.innerHTML = '<span class="toast-icon"><i class="' + (icons[type] || icons.info) + '"></i></span>' + message;
        container.appendChild(toast);
        setTimeout(function () { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 4000);
    }

    form.addEventListener('submit', async function (e) {
        e.preventDefault();

        if (!form.checkValidity()) {
            form.reportValidity();
            return;
        }

        submitBtn.disabled = true;
        submitBtn.classList.add('loading');

        const licenseTypes = {
            trial: '30-Day Free Trial',
            '6mo': '6 Months',
            '12mo': '12 Months',
            lifetime: 'Lifetime'
        };
        const licenseTypeVal = document.getElementById('licenseType').value;
        const licenseTypeLabel = licenseTypes[licenseTypeVal] || licenseTypeVal;

        const productVal = document.getElementById('product').value;

        const data = {
            _subject: '[TRIAL] ' + productVal.toUpperCase() + ' - ' + licenseTypeLabel + ' - ' + document.getElementById('email').value.trim(),
            name: document.getElementById('name').value.trim(),
            email: document.getElementById('email').value.trim(),
            licenseType: licenseTypeLabel,
            product: document.getElementById('product').value,
            purpose: document.getElementById('purpose').value,
            officeVersion: document.getElementById('officeVersion').value,
            notes: document.getElementById('notes').value.trim(),
        };

        try {
            const resp = await fetch('https://formsubmit.co/ajax/apnan2010@gmail.com', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                body: JSON.stringify(data)
            });

            if (!resp.ok) throw new Error('Server error');

            showSuccess();
        } catch (err) {
            submitBtn.disabled = false;
            submitBtn.classList.remove('loading');
            showToast('Something went wrong. Please try again or email us at apnan2010@gmail.com', 'error');
        }
    });

    function showSuccess() {
        const container = document.getElementById('formContainer');
        container.innerHTML =
            '<div class="success-card">' +
            '<h3>✓ Request Submitted</h3>' +
            '<p>Thank you! Your license key request has been received.<br><br>' +
            'We will process your request and send the license key to your email within 24 hours.<br><br>' +
            'If you don\'t receive the email, please check your spam folder or <a href="mailto:apnan2010@gmail.com" style="color:#e94560;">contact us</a>.</p>' +
            '<br><a href="index.html" class="btn" style="display:inline-block;margin-top:16px;">Back to Home</a>' +
            '</div>';
    }
})();
