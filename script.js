// ==================== CONFIG ====================
let CONFIG = {
    operating247: true,
    customHours: "24/7 Available",
    whatsappNumber: "919448301456",
    businessName: "Hitendra Tours & Travels"
};
const GAS_URL = "https://script.google.com/macros/s/AKfycbzupS81FcWUplphlki0G-wTtvnhhfaacXk1_qoxRqkRDXHLTYIN81dL88e8DpdalnycTA/exec";

// Anti‑bot: record form start time
let formStartTime = Date.now();

// ==================== INIT ====================
document.addEventListener('DOMContentLoaded', () => {
    loadConfig();
    setupDatePicker();
    setupTripTypeToggle();
    setupMultiStop();
    setupFormSubmit();
    removeLoadingOverlay();
});

function removeLoadingOverlay() {
    setTimeout(() => {
        const overlay = document.getElementById('loadingOverlay');
        if (overlay) {
            overlay.style.opacity = '0';
            setTimeout(() => overlay.style.display = 'none', 500);
        }
    }, 500);
}

async function loadConfig() {
    try {
        const res = await fetch(`${GAS_URL}?action=getConfig`);
        if (res.ok) {
            const data = await res.json();
            if (data.success) {
                CONFIG = { ...CONFIG, ...data.config };
                updateStatusDisplay();
                updateFooterHours();
            }
        }
    } catch(e) { console.log(e); }
}

function updateStatusDisplay() {
    const dot = document.getElementById('statusDot');
    const text = document.getElementById('statusText');
    const hours = document.getElementById('hoursText');
    if (CONFIG.operating247) {
        dot.classList.remove('closed');
        text.textContent = '🟢 We are open 24/7';
        if (hours) hours.textContent = 'Available anytime, day or night';
    } else {
        dot.classList.add('closed');
        text.textContent = '🔴 Limited Hours';
        if (hours) hours.textContent = CONFIG.customHours;
    }
}

function updateFooterHours() {
    const el = document.getElementById('footerHours');
    if (el) el.innerHTML = CONFIG.operating247 ? '🟢 24/7 Available<br>Call anytime' : CONFIG.customHours;
}

function setupDatePicker() {
    const dateInput = document.getElementById('journeyDate');
    if (dateInput) dateInput.min = new Date().toISOString().split('T')[0];
}

// ==================== TRIP TYPE TOGGLE ====================
function setupTripTypeToggle() {
    const tripType = document.getElementById('tripType');
    const roundFields = document.getElementById('roundTripFields');
    const multiFields = document.getElementById('multiStopFields');
    const dropLabel = document.getElementById('dropLabel');
    tripType.addEventListener('change', () => {
        const val = tripType.value;
        roundFields.style.display = val === 'roundtrip' ? 'block' : 'none';
        multiFields.style.display = val === 'multistop' ? 'block' : 'none';
        dropLabel.innerText = val === 'multistop' ? 'Final Drop Location *' : 'Drop Location *';
        if (val === 'roundtrip') {
            document.getElementById('returnDate').required = true;
            document.getElementById('returnTime').required = true;
        } else {
            document.getElementById('returnDate').required = false;
            document.getElementById('returnTime').required = false;
        }
    });
}

// ==================== MULTI‑STOP ====================
let stopCount = 1;
const MAX_STOPS = 3;

function setupMultiStop() {
    const addBtn = document.getElementById('addStopBtn');
    if (addBtn) addBtn.addEventListener('click', addStopField);
}

function addStopField() {
    if (stopCount >= MAX_STOPS) {
        showToast(`Maximum ${MAX_STOPS} stops allowed`, 'error');
        return;
    }
    stopCount++;
    const container = document.getElementById('stopsContainer');
    const div = document.createElement('div');
    div.className = 'stop-entry';
    div.innerHTML = `
        <div class="form-row">
            <div class="form-group"><label>Stop ${stopCount} Location</label><input type="text" class="stop-location" placeholder="e.g., Gateway of India"></div>
            <div class="form-group"><label>Halt (minutes)</label><input type="number" class="stop-halt" placeholder="30" min="0" step="10"></div>
        </div>
        <button type="button" class="remove-stop-btn" onclick="this.parentElement.remove(); stopCount--;">Remove stop</button>
    `;
    container.appendChild(div);
}

// ==================== FORM HANDLING ====================
function setupFormSubmit() {
    document.getElementById('bookingForm').addEventListener('submit', handleFormSubmit);
}

function collectFormData() {
    const tripType = document.getElementById('tripType').value;
    const stops = [];
    if (tripType === 'multistop') {
        const locations = document.querySelectorAll('.stop-location');
        const halts = document.querySelectorAll('.stop-halt');
        for (let i = 0; i < locations.length; i++) {
            if (locations[i].value.trim()) {
                stops.push({ location: locations[i].value.trim(), halt: halts[i].value ? parseInt(halts[i].value) : 0 });
            }
        }
    }
    return {
        bookingId: generateBookingId(),
        fullName: document.getElementById('fullName').value.trim(),
        phone: document.getElementById('phone').value.trim(),
        email: document.getElementById('email').value.trim(),
        serviceType: document.getElementById('serviceType').value,
        tripType: tripType,
        pickup: document.getElementById('pickup').value.trim(),
        drop: document.getElementById('drop').value.trim(),
        returnDate: document.getElementById('returnDate').value,
        returnTime: document.getElementById('returnTime').value,
        returnDrop: document.getElementById('returnDrop').value.trim(),
        stops: stops,
        landmark: document.getElementById('landmark').value.trim(),
        journeyDate: document.getElementById('journeyDate').value,
        pickupTime: document.getElementById('pickupTime').value,
        passengers: document.getElementById('passengers').value,
        luggage: document.getElementById('luggage').value,
        timestamp: new Date().toISOString(),
        status: 'Pending'
    };
}

function validateForm() {
    // Honeypot
    if (document.getElementById('website').value !== "") {
        showToast("Spam detected. Please try again.", "error");
        return false;
    }
    // Time validation (minimum 5 seconds)
    const timeTaken = (Date.now() - formStartTime) / 1000;
    if (timeTaken < 5) {
        showToast("Please take a moment to fill the form properly.", "error");
        return false;
    }
    // Required fields
    const required = ['fullName', 'phone', 'email', 'serviceType', 'pickup', 'drop', 'journeyDate', 'pickupTime', 'passengers'];
    for (let f of required) {
        const el = document.getElementById(f);
        if (!el.value.trim()) {
            showToast(`Please fill ${el.previousElementSibling.innerText}`, 'error');
            el.focus();
            return false;
        }
    }
    const tripType = document.getElementById('tripType').value;
    if (tripType === 'roundtrip') {
        if (!document.getElementById('returnDate').value || !document.getElementById('returnTime').value) {
            showToast('Please provide return date and time for round trip', 'error');
            return false;
        }
    }
    if (tripType === 'multistop') {
        const stopLocs = document.querySelectorAll('.stop-location');
        let hasValid = false;
        for (let i = 0; i < stopLocs.length; i++) {
            if (stopLocs[i].value.trim()) hasValid = true;
        }
        if (!hasValid && stopLocs.length > 0) {
            showToast('Please add at least one valid stop location', 'error');
            return false;
        }
    }
    const phone = document.getElementById('phone').value;
    if (!/^\d{10}$/.test(phone)) { showToast('Invalid 10-digit mobile number', 'error'); return false; }
    const email = document.getElementById('email').value;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { showToast('Invalid email', 'error'); return false; }
    return true;
}

function generateBookingId() {
    const prefix = 'HIT';
    const d = new Date();
    const ts = d.getFullYear().toString().slice(-2) + (d.getMonth()+1).toString().padStart(2,'0') + d.getDate().toString().padStart(2,'0') + d.getHours().toString().padStart(2,'0') + d.getMinutes().toString().padStart(2,'0') + d.getSeconds().toString().padStart(2,'0');
    const rand = Math.floor(Math.random()*1000).toString().padStart(3,'0');
    return `${prefix}${ts}${rand}`;
}

// ==================== VEHICLE AVAILABILITY ====================
async function checkVehicleAvailability(date, time) {
    try {
        const res = await fetch(`${GAS_URL}?action=checkAvailability&date=${date}&time=${time}`);
        return await res.json();
    } catch(e) { return { available: true, availableCount: 2, message: 'Available' }; }
}

async function handleFormSubmit(e) {
    e.preventDefault();
    if (!validateForm()) return;
    const bookingData = collectFormData();
    showToast('Checking vehicle availability...', 'success');
    const availability = await checkVehicleAvailability(bookingData.journeyDate, bookingData.pickupTime);
    if (!availability.available) {
        showNoVehiclesModal(availability.message);
        return;
    }
    bookingData.availableCount = availability.availableCount;
    showBookingSummary(bookingData);
}

function showNoVehiclesModal(message) {
    let modal = document.getElementById('noVehiclesModal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'noVehiclesModal'; modal.className = 'modal-overlay'; document.body.appendChild(modal); }
    modal.innerHTML = `
        <div class="modal">
            <div class="modal-header" style="background: #e74c3c;"><h3><i class="fas fa-truck"></i> 🚐 Vehicles Fully Booked</h3><button class="modal-close" onclick="closeNoVehiclesModal()">&times;</button></div>
            <div class="modal-body" style="text-align:center;">
                <i class="fas fa-clock" style="font-size:60px; color:#e74c3c; margin-bottom:20px;"></i>
                <p style="font-size:18px; margin-bottom:15px;">${message}</p>
                <div style="background:#fff3cd; padding:15px; border-radius:8px; margin:20px 0;"><strong>💡 Suggestions:</strong><br>• Try a different time slot<br>• Choose another date<br>• Call us directly at <strong>94483 01456</strong></div>
                <button onclick="closeNoVehiclesModal()" class="submit-btn" style="background:#1e3c5c;">Choose Different Time</button>
            </div>
        </div>
    `;
    setTimeout(() => modal.classList.add('active'), 10);
}

function closeNoVehiclesModal() {
    const modal = document.getElementById('noVehiclesModal');
    if (modal) modal.classList.remove('active');
}

// ==================== BOOKING SUMMARY & WHATSAPP ====================
function showBookingSummary(bookingData) {
    let modal = document.getElementById('summaryModal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'summaryModal'; modal.className = 'modal-overlay'; document.body.appendChild(modal); }
    const formattedDate = new Date(bookingData.journeyDate).toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' });
    let stopsHtml = '';
    if (bookingData.stops && bookingData.stops.length) {
        stopsHtml = '<div class="summary-item"><strong>Stops:</strong><span>' + bookingData.stops.map(s => `${s.location} (${s.halt} min)`).join(', ') + '</span></div>';
    }
    modal.innerHTML = `
        <div class="modal">
            <div class="modal-header"><h3><i class="fas fa-file-alt"></i> Booking Summary</h3><button class="modal-close" onclick="closeModal()">&times;</button></div>
            <div class="modal-body">
                <div class="booking-summary">
                    <div class="summary-item"><strong>Booking ID:</strong> <span>${bookingData.bookingId}</span></div>
                    <div class="summary-item"><strong>Name:</strong> <span>${escapeHtml(bookingData.fullName)}</span></div>
                    <div class="summary-item"><strong>Phone:</strong> <span>${bookingData.phone}</span></div>
                    <div class="summary-item"><strong>Email:</strong> <span>${escapeHtml(bookingData.email)}</span></div>
                    <div class="summary-item"><strong>Service:</strong> <span>${bookingData.serviceType}</span></div>
                    <div class="summary-item"><strong>Trip Type:</strong> <span>${bookingData.tripType === 'oneway' ? 'One Way' : (bookingData.tripType === 'roundtrip' ? 'Round Trip' : 'Multi‑Stop')}</span></div>
                    <div class="summary-item"><strong>Pickup:</strong> <span>${escapeHtml(bookingData.pickup)}</span></div>
                    ${stopsHtml}
                    <div class="summary-item"><strong>Drop:</strong> <span>${escapeHtml(bookingData.drop)}</span></div>
                    ${bookingData.tripType === 'roundtrip' ? `<div class="summary-item"><strong>Return:</strong> <span>${bookingData.returnDate} ${bookingData.returnTime} ${bookingData.returnDrop ? '→ '+bookingData.returnDrop : ''}</span></div>` : ''}
                    <div class="summary-item"><strong>Date:</strong> <span>${formattedDate}</span></div>
                    <div class="summary-item"><strong>Time:</strong> <span>${bookingData.pickupTime}</span></div>
                    <div class="summary-item"><strong>Passengers:</strong> <span>${bookingData.passengers}</span></div>
                </div>
                <div class="modal-actions">
                    <button class="btn-cancel" onclick="closeModal()">Edit Details</button>
                    <button class="btn-whatsapp" onclick="sendToWhatsApp('${bookingData.bookingId}')"><i class="fab fa-whatsapp"></i> Send via WhatsApp</button>
                </div>
            </div>
        </div>
    `;
    window.currentBookingData = bookingData;
    setTimeout(() => modal.classList.add('active'), 10);
}

function createWhatsAppMessage(data) {
    const dateObj = new Date(data.journeyDate);
    const formattedDate = dateObj.toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    let msg = `🚐 *NEW BOOKING REQUEST - ${data.bookingId}* 🚐\n\n`;
    msg += `*Customer:* ${data.fullName}\n📞 ${data.phone}\n✉️ ${data.email}\n\n`;
    msg += `*Trip Type:* ${data.tripType === 'oneway' ? 'One Way' : (data.tripType === 'roundtrip' ? 'Round Trip' : 'Multi‑Stop')}\n`;
    msg += `*Service:* ${data.serviceType}\n`;
    msg += `*Pickup:* ${data.pickup}\n`;
    if (data.stops && data.stops.length) {
        msg += `*Intermediate Stops:*\n`;
        data.stops.forEach((s, idx) => { msg += `   ${idx+1}. ${s.location} (halt ${s.halt} min)\n`; });
    }
    msg += `*Final Drop:* ${data.drop}\n`;
    if (data.tripType === 'roundtrip') {
        const retDate = new Date(data.returnDate).toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
        msg += `*Return Date:* ${retDate}\n*Return Pickup Time:* ${data.returnTime}\n`;
        msg += `*Return Drop:* ${data.returnDrop || 'Same as original pickup'}\n`;
    }
    if (data.landmark) msg += `*Landmark:* ${data.landmark}\n`;
    msg += `*Date:* ${formattedDate}\n*Time:* ${data.pickupTime}\n`;
    msg += `*Passengers:* ${data.passengers}\n*Luggage:* ${data.luggage}\n\n`;
    msg += `*Service Highlights:*\n✅ Free flight tracking\n✅ Professional chauffeur\n✅ Sanitized EECO Van\n✅ Real-time WhatsApp updates\n\n`;
    msg += `_Please share fare quote and confirm availability._`;
    return msg;
}

async function sendToWhatsApp(bookingId) {
    const data = window.currentBookingData;
    if (!data) return;
    const message = createWhatsAppMessage(data);
    const encoded = encodeURIComponent(message);
    const url = `https://wa.me/${CONFIG.whatsappNumber}?text=${encoded}`;
    saveBookingToSheet(data);
    if (document.getElementById('emailCopy').checked) sendEmailCopy(data);
    closeModal();
    showToast('Redirecting to WhatsApp...', 'success');
    setTimeout(() => window.open(url, '_blank'), 500);
}

async function saveBookingToSheet(data) {
    try {
        await fetch(GAS_URL, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'saveBooking', data }) });
    } catch(e) { console.log(e); }
}

async function sendEmailCopy(data) {
    try {
        await fetch(`${GAS_URL}?action=sendEmailCopy`, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
        showToast('Confirmation email sent!', 'success');
    } catch(e) { console.log(e); }
}

function closeModal() {
    const modal = document.getElementById('summaryModal');
    if (modal) modal.classList.remove('active');
}

function showToast(msg, type = 'success') {
    let toast = document.getElementById('toast');
    if (!toast) { toast = document.createElement('div'); toast.id = 'toast'; toast.className = 'toast'; document.body.appendChild(toast); }
    toast.textContent = msg;
    toast.className = `toast ${type} show`;
    setTimeout(() => toast.classList.remove('show'), 3000);
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// Global functions for modal buttons
window.closeModal = closeModal;
window.sendToWhatsApp = sendToWhatsApp;
window.closeNoVehiclesModal = closeNoVehiclesModal;