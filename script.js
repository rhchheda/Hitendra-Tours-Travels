// Configuration - Default values (overridden by Google Sheets)
let CONFIG = {
    operating247: true,
    customHours: "24/7 Available",
    whatsappNumber: "919448301456",
    businessName: "Hitendra Tours & Travels"
};

// Google Apps Script URL - REPLACE WITH YOUR DEPLOYED URL
const GAS_URL = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec";

// Initialize when DOM loads
document.addEventListener('DOMContentLoaded', function() {
    loadConfig();
    setupDatePicker();
    setupFormSubmit();
    removeLoadingOverlay();
});

// Remove loading overlay
function removeLoadingOverlay() {
    setTimeout(() => {
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.style.opacity = '0';
            setTimeout(() => {
                overlay.style.display = 'none';
            }, 500);
        }
    }, 500);
}

// Load configuration from Google Sheets
async function loadConfig() {
    try {
        const response = await fetch(`${GAS_URL}?action=getConfig`);
        if (response.ok) {
            const data = await response.json();
            if (data.success) {
                CONFIG = { ...CONFIG, ...data.config };
                updateStatusDisplay();
                updateFooterHours();
            }
        }
    } catch (error) {
        console.log('Using default config:', error);
        updateStatusDisplay();
        updateFooterHours();
    }
}

// Update status display based on config
function updateStatusDisplay() {
    const statusDot = document.getElementById('statusDot');
    const statusText = document.getElementById('statusText');
    const hoursText = document.getElementById('hoursText');
    
    if (CONFIG.operating247) {
        statusDot.classList.remove('closed');
        statusText.textContent = '🟢 We are open 24/7';
        if (hoursText) hoursText.textContent = 'Available anytime, day or night';
    } else {
        statusDot.classList.add('closed');
        statusText.textContent = '🔴 Limited Hours';
        if (hoursText) hoursText.textContent = CONFIG.customHours;
    }
}

// Update footer hours
function updateFooterHours() {
    const footerHours = document.getElementById('footerHours');
    if (footerHours) {
        if (CONFIG.operating247) {
            footerHours.innerHTML = '🟢 24/7 Available<br>Call anytime';
        } else {
            footerHours.innerHTML = CONFIG.customHours;
        }
    }
}

// Setup date picker (disable past dates)
function setupDatePicker() {
    const dateInput = document.getElementById('journeyDate');
    if (dateInput) {
        const today = new Date().toISOString().split('T')[0];
        dateInput.min = today;
    }
}

// Setup form submit
function setupFormSubmit() {
    const form = document.getElementById('bookingForm');
    if (form) {
        form.addEventListener('submit', handleFormSubmit);
    }
}

// Check vehicle availability before submission
async function checkVehicleAvailability(date, time) {
    try {
        const response = await fetch(`${GAS_URL}?action=checkAvailability&date=${date}&time=${time}`);
        const data = await response.json();
        return data;
    } catch (error) {
        console.log('Availability check failed, assuming available:', error);
        return { available: true, availableCount: 2, message: 'Available' };
    }
}

// Handle form submission with availability check
async function handleFormSubmit(e) {
    e.preventDefault();
    
    if (!validateForm()) return;
    
    const bookingData = collectFormData();
    
    // Show loading
    showToast('Checking vehicle availability...', 'success');
    
    // Check availability
    const availability = await checkVehicleAvailability(bookingData.journeyDate, bookingData.pickupTime);
    
    if (!availability.available) {
        showNoVehiclesModal(availability.message);
        return;
    }
    
    bookingData.availableCount = availability.availableCount;
    showBookingSummary(bookingData);
}

// Show "No Vehicles Available" modal
function showNoVehiclesModal(message) {
    let modal = document.getElementById('noVehiclesModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'noVehiclesModal';
        modal.className = 'modal-overlay';
        document.body.appendChild(modal);
    }
    
    modal.innerHTML = `
        <div class="modal">
            <div class="modal-header" style="background: #e74c3c;">
                <h3><i class="fas fa-truck"></i> 🚐 Vehicles Fully Booked</h3>
                <button class="modal-close" onclick="closeNoVehiclesModal()">&times;</button>
            </div>
            <div class="modal-body" style="text-align: center;">
                <i class="fas fa-clock" style="font-size: 60px; color: #e74c3c; margin-bottom: 20px;"></i>
                <p style="font-size: 18px; margin-bottom: 15px;">${message}</p>
                <div style="background: #fff3cd; padding: 15px; border-radius: 8px; margin: 20px 0;">
                    <strong>💡 Suggestions:</strong><br>
                    • Try a different time slot<br>
                    • Choose another date<br>
                    • Call us directly at <strong>94483 01456</strong>
                </div>
                <button onclick="closeNoVehiclesModal()" class="submit-btn" style="background: #1e3c5c; margin-top: 10px;">
                    Choose Different Time
                </button>
            </div>
        </div>
    `;
    
    setTimeout(() => {
        modal.classList.add('active');
    }, 10);
}

function closeNoVehiclesModal() {
    const modal = document.getElementById('noVehiclesModal');
    if (modal) {
        modal.classList.remove('active');
    }
}

// Validate form
function validateForm() {
    const required = ['fullName', 'phone', 'email', 'serviceType', 'pickup', 'drop', 'journeyDate', 'pickupTime', 'passengers'];
    
    for (const field of required) {
        const input = document.getElementById(field);
        if (!input || !input.value.trim()) {
            showToast(`Please fill ${input?.previousElementSibling?.innerText || field}`, 'error');
            input?.focus();
            return false;
        }
    }
    
    const phone = document.getElementById('phone').value;
    if (!/^\d{10}$/.test(phone)) {
        showToast('Please enter a valid 10-digit mobile number', 'error');
        return false;
    }
    
    const email = document.getElementById('email').value;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        showToast('Please enter a valid email address', 'error');
        return false;
    }
    
    return true;
}

// Collect form data
function collectFormData() {
    return {
        bookingId: generateBookingId(),
        fullName: document.getElementById('fullName').value.trim(),
        phone: document.getElementById('phone').value.trim(),
        email: document.getElementById('email').value.trim(),
        serviceType: document.getElementById('serviceType').value,
        pickup: document.getElementById('pickup').value.trim(),
        drop: document.getElementById('drop').value.trim(),
        landmark: document.getElementById('landmark').value.trim(),
        journeyDate: document.getElementById('journeyDate').value,
        pickupTime: document.getElementById('pickupTime').value,
        passengers: document.getElementById('passengers').value,
        luggage: document.getElementById('luggage').value,
        timestamp: new Date().toISOString(),
        status: 'Pending'
    };
}

// Generate unique booking ID
function generateBookingId() {
    const prefix = 'HIT';
    const date = new Date();
    const timestamp = date.getFullYear().toString().slice(-2) + 
                    (date.getMonth() + 1).toString().padStart(2, '0') + 
                    date.getDate().toString().padStart(2, '0') +
                    date.getHours().toString().padStart(2, '0') +
                    date.getMinutes().toString().padStart(2, '0') +
                    date.getSeconds().toString().padStart(2, '0');
    const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return `${prefix}${timestamp}${random}`;
}

// Show booking summary modal
function showBookingSummary(bookingData) {
    let modal = document.getElementById('summaryModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'summaryModal';
        modal.className = 'modal-overlay';
        document.body.appendChild(modal);
    }
    
    const formattedDate = new Date(bookingData.journeyDate).toLocaleDateString('en-IN', {
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    });
    
    modal.innerHTML = `
        <div class="modal">
            <div class="modal-header">
                <h3><i class="fas fa-file-alt"></i> Booking Summary</h3>
                <button class="modal-close" onclick="closeModal()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="booking-summary">
                    <div class="summary-item"><strong>Booking ID:</strong> <span>${bookingData.bookingId}</span></div>
                    <div class="summary-item"><strong>Name:</strong> <span>${escapeHtml(bookingData.fullName)}</span></div>
                    <div class="summary-item"><strong>Phone:</strong> <span>${bookingData.phone}</span></div>
                    <div class="summary-item"><strong>Email:</strong> <span>${escapeHtml(bookingData.email)}</span></div>
                    <div class="summary-item"><strong>Service:</strong> <span>${bookingData.serviceType}</span></div>
                    <div class="summary-item"><strong>Pickup:</strong> <span>${escapeHtml(bookingData.pickup)}</span></div>
                    <div class="summary-item"><strong>Drop:</strong> <span>${escapeHtml(bookingData.drop)}</span></div>
                    ${bookingData.landmark ? `<div class="summary-item"><strong>Landmark:</strong> <span>${escapeHtml(bookingData.landmark)}</span></div>` : ''}
                    <div class="summary-item"><strong>Date:</strong> <span>${formattedDate}</span></div>
                    <div class="summary-item"><strong>Time:</strong> <span>${bookingData.pickupTime}</span></div>
                    <div class="summary-item"><strong>Passengers:</strong> <span>${bookingData.passengers}</span></div>
                    <div class="summary-item"><strong>Luggage:</strong> <span>${bookingData.luggage}</span></div>
                </div>
                <p><i class="fas fa-info-circle"></i> Click "Send via WhatsApp" to complete your booking request.</p>
                <div class="modal-actions">
                    <button class="btn-cancel" onclick="closeModal()">Edit Details</button>
                    <button class="btn-whatsapp" onclick="sendToWhatsApp('${bookingData.bookingId}')">
                        <i class="fab fa-whatsapp"></i> Send via WhatsApp
                    </button>
                </div>
            </div>
        </div>
    `;
    
    window.currentBookingData = bookingData;
    
    setTimeout(() => {
        modal.classList.add('active');
    }, 10);
}

// Send to WhatsApp
async function sendToWhatsApp(bookingId) {
    const bookingData = window.currentBookingData;
    if (!bookingData) return;
    
    const message = createWhatsAppMessage(bookingData);
    const encodedMessage = encodeURIComponent(message);
    const whatsappUrl = `https://wa.me/${CONFIG.whatsappNumber}?text=${encodedMessage}`;
    
    saveBookingToSheet(bookingData);
    
    const emailCopy = document.getElementById('emailCopy').checked;
    if (emailCopy) {
        sendEmailCopy(bookingData);
    }
    
    closeModal();
    showToast('Redirecting to WhatsApp...', 'success');
    
    setTimeout(() => {
        window.open(whatsappUrl, '_blank');
    }, 500);
}

// Enhanced WhatsApp message with service highlights
function createWhatsAppMessage(data) {
    const dateObj = new Date(data.journeyDate);
    const formattedDate = dateObj.toLocaleDateString('en-IN', {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
    });
    
    let message = `🚐 *NEW BOOKING REQUEST - ${data.bookingId}* 🚐\n\n`;
    message += `*Customer Details:*\n`;
    message += `👤 Name: ${data.fullName}\n`;
    message += `📞 Phone: ${data.phone}\n`;
    message += `✉️ Email: ${data.email}\n\n`;
    message += `*Trip Details:*\n`;
    message += `🛎️ Service: ${data.serviceType}\n`;
    message += `📍 Pickup: ${data.pickup}\n`;
    message += `🏁 Drop: ${data.drop}\n`;
    if (data.landmark) message += `📌 Landmark: ${data.landmark}\n`;
    message += `📅 Date: ${formattedDate}\n`;
    message += `⏰ Time: ${data.pickupTime}\n`;
    message += `👥 Passengers: ${data.passengers}\n`;
    message += `🧳 Luggage: ${data.luggage}\n\n`;
    message += `*Service Highlights:*\n`;
    message += `✅ Free flight tracking\n`;
    message += `✅ Professional chauffeur\n`;
    message += `✅ Sanitized EECO Van\n`;
    message += `✅ Real-time WhatsApp updates\n\n`;
    message += `_Please share fare quote and confirm availability._\n`;
    message += `Reply to this message to connect with customer.`;
    
    return message;
}

// Save booking to Google Sheets
async function saveBookingToSheet(bookingData) {
    try {
        await fetch(GAS_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'saveBooking', data: bookingData })
        });
        console.log('Booking saved to sheet');
    } catch (error) {
        console.log('Could not save to sheet:', error);
    }
}

// Send email copy using GAS
async function sendEmailCopy(bookingData) {
    try {
        await fetch(`${GAS_URL}?action=sendEmailCopy`, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(bookingData)
        });
        showToast('Confirmation email sent!', 'success');
    } catch (error) {
        console.log('Email send error:', error);
    }
}

function closeModal() {
    const modal = document.getElementById('summaryModal');
    if (modal) modal.classList.remove('active');
}

function showToast(message, type = 'success') {
    let toast = document.getElementById('toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'toast';
        toast.className = 'toast';
        document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.className = `toast ${type} show`;
    setTimeout(() => toast.classList.remove('show'), 3000);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Global functions for modals
window.closeModal = closeModal;
window.sendToWhatsApp = sendToWhatsApp;
window.closeNoVehiclesModal = closeNoVehiclesModal;