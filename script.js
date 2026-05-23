// ==================== CONFIGURATION ====================
let CONFIG = {
    operating247: true,
    customHours: "24/7 Available",
    whatsappNumber: "919448301456",
    businessName: "Hitendra Tours & Travels"
};
const GAS_URL = "https://script.google.com/macros/s/AKfycbwFls5uasexKXAPPkRrjJ6xwLML-yS0u5GnsHwr3YxOxItT6fg16JPR6XT7IL3Ui5t6/exec";
let formStartTime = Date.now();

// ==================== HELPER: Format Date in IST (DD-MMM-YYYY) ====================
function formatDateIST(dateStr) {
    // dateStr is YYYY-MM-DD from input
    const [year, month, day] = dateStr.split('-');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${day}-${months[parseInt(month)-1]}-${year}`;
}

// ==================== HELPER: Format Time in IST (hh:mm AM/PM) ====================
function formatTimeIST(time24) {
    // time24 is "HH:MM"
    let [hours, minutes] = time24.split(':');
    hours = parseInt(hours);
    const ampm = hours >= 12 ? 'PM' : 'AM';
    let hour12 = hours % 12;
    if (hour12 === 0) hour12 = 12;
    return `${hour12}:${minutes} ${ampm}`;
}

// ==================== HELPER: Format Full DateTime in IST for display ====================
function formatDateTimeIST(dateStr, timeStr) {
    return `${formatDateIST(dateStr)} at ${formatTimeIST(timeStr)}`;
}

// ==================== UPDATE DISPLAYS (Hero, Footer) ====================
// (unchanged – they use config from sheet)

// ==================== TRIP TYPE & MULTI-STOP (unchanged) ====================

// ==================== FORM SUBMISSION & VALIDATION ====================
function validateForm() {
    // Honeypot, time validation, required fields (same as before)
    // ... (keep existing validation)
}

function collectFormData() {
    // same as before
}

function generateBookingId() {
    // same as before
}

// ==================== CREATE WHATSAPP MESSAGE (uses IST formatted dates) ====================
function createWhatsAppMessage(data) {
    const journeyDate = data.journeyDate;
    const pickupTime = data.pickupTime;
    const formattedJourney = formatDateTimeIST(journeyDate, pickupTime);
    
    let msg = `*NEW BOOKING REQUEST - ${data.bookingId}*\n\n`;
    msg += `Customer: ${data.fullName}\nPhone: ${data.phone}\nEmail: ${data.email}\n\n`;
    msg += `Trip: ${data.tripType === 'oneway' ? 'One Way' : (data.tripType === 'roundtrip' ? 'Round Trip' : 'Multi-Stop')}\n`;
    msg += `Service: ${data.serviceType}\nPickup: ${data.pickup}\n`;
    if (data.stops && data.stops.length) {
        msg += `Stops:\n`;
        data.stops.forEach((s, idx) => { msg += `  ${idx+1}. ${s.location} (halt ${s.halt} min)\n`; });
    }
    msg += `Drop: ${data.drop}\n`;
    if (data.tripType === 'roundtrip') {
        const returnDateTime = formatDateTimeIST(data.returnDate, data.returnTime);
        msg += `Return: ${returnDateTime}\nReturn Drop: ${data.returnDrop || 'Same as pickup'}\n`;
    }
    if (data.landmark) msg += `Landmark: ${data.landmark}\n`;
    msg += `Journey: ${formattedJourney}\n`;
    msg += `Passengers: ${data.passengers}\nLuggage: ${data.luggage}\n\n`;
    msg += `Service Highlights: Free flight tracking, Professional chauffeur, Sanitized EECO Van\n\n`;
    msg += `To confirm, please reply with "OK".`;
    return msg;
}

// ==================== POST-SUBMIT MODAL (uses IST) ====================
function showPostSubmitModal(data) {
    // Replace the call to toLocaleDateString with formatDateIST
    const formattedDate = formatDateIST(data.journeyDate);
    // ... rest of modal HTML (use formattedDate)
}

// ==================== SEND EMAIL COPY (frontend call, actual email sent by GAS) ====================
async function sendEmailCopy(data) {
    try {
        await fetch(`${GAS_URL}?action=sendEmailCopy`, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        showToast('Confirmation email sent!', 'success');
    } catch(e) { console.log(e); }
}
// ==================== VEHICLE AVAILABILITY ====================
async function checkVehicleAvailability(date, time) {
    try {
        let res = await fetch(`${GAS_URL}?action=checkAvailability&date=${date}&time=${time}`);
        return await res.json();
    } catch (e) { return { available: true, availableCount: 2, message: 'Available' }; }
}

async function saveBookingToSheet(data) {
    try {
        await fetch(GAS_URL, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'saveBooking', data: data }) });
        return { success: true };
    } catch (e) { console.log(e); return { success: false }; }
}

async function handleFormSubmit(e) {
    e.preventDefault();
    if (!validateForm()) return;
    let bookingData = collectFormData();
    showToast('Checking vehicle availability...', 'success');
    let avail = await checkVehicleAvailability(bookingData.journeyDate, bookingData.pickupTime);
    if (!avail.available) { showNoVehiclesModal(avail.message); return; }
    showToast('Saving your booking...', 'success');
    let saved = await saveBookingToSheet(bookingData);
    if (!saved.success) { showToast('Failed to save booking. Please try again.', 'error'); return; }
    window.currentBookingData = bookingData;
    showPostSubmitModal(bookingData);
}

function showNoVehiclesModal(msg) {
    let modal = document.getElementById('noVehiclesModal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'noVehiclesModal'; modal.className = 'modal-overlay'; document.body.appendChild(modal); }
    modal.innerHTML = `<div class="modal"><div class="modal-header" style="background:#e74c3c;"><h3><i class="fas fa-truck"></i> 🚐 Vehicles Fully Booked</h3><button class="modal-close" onclick="closeNoVehiclesModal()">&times;</button></div><div class="modal-body" style="text-align:center;"><i class="fas fa-clock" style="font-size:60px; color:#e74c3c; margin-bottom:20px;"></i><p style="font-size:18px; margin-bottom:15px;">${msg}</p><div style="background:#fff3cd; padding:15px; border-radius:8px; margin:20px 0;"><strong>💡 Suggestions:</strong><br>• Try a different time slot<br>• Choose another date<br>• Call us directly at <strong>94483 01456</strong></div><button onclick="closeNoVehiclesModal()" class="submit-btn" style="background:#1e3c5c;">Choose Different Time</button></div></div>`;
    setTimeout(() => modal.classList.add('active'), 10);
}

function closeNoVehiclesModal() {
    let modal = document.getElementById('noVehiclesModal');
    if (modal) modal.classList.remove('active');
}

// ==================== POST-SUBMIT MODAL ====================
function showPostSubmitModal(data) {
    let modal = document.getElementById('postSubmitModal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'postSubmitModal'; modal.className = 'modal-overlay'; document.body.appendChild(modal); }
    let formattedDate = new Date(data.journeyDate).toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' });
    modal.innerHTML = `<div class="modal"><div class="modal-header" style="background:#27ae60;"><h3><i class="fas fa-check-circle"></i> Booking Submitted Successfully!</h3><button class="modal-close" onclick="closePostSubmitModal()">&times;</button></div><div class="modal-body" style="text-align:center;"><i class="fas fa-save" style="font-size:48px; color:#27ae60; margin-bottom:15px;"></i><p>Your booking has been saved. Booking ID: <strong>${data.bookingId}</strong></p><p>Would you like to send the details via WhatsApp to complete the request?</p><div class="modal-actions" style="justify-content:center; margin-top:20px;"><button class="btn-cancel" onclick="closePostSubmitModal()">Close</button><button class="btn-whatsapp" onclick="sendWhatsAppFromModal()"><i class="fab fa-whatsapp"></i> Send via WhatsApp</button></div></div></div>`;
    setTimeout(() => modal.classList.add('active'), 10);
}

function closePostSubmitModal() {
    let modal = document.getElementById('postSubmitModal');
    if (modal) modal.classList.remove('active');
}

function createWhatsAppMessage(data) {
    let d = new Date(data.journeyDate);
    let formattedDate = d.toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    let msg = `*NEW BOOKING REQUEST - ${data.bookingId}*\n\n`;
    msg += `Customer: ${data.fullName}\nPhone: ${data.phone}\nEmail: ${data.email}\n\n`;
    msg += `Trip: ${data.tripType === 'oneway' ? 'One Way' : (data.tripType === 'roundtrip' ? 'Round Trip' : 'Multi-Stop')}\n`;
    msg += `Service: ${data.serviceType}\nPickup: ${data.pickup}\n`;
    if (data.stops && data.stops.length) {
        msg += `Stops:\n`;
        data.stops.forEach((s, idx) => { msg += `  ${idx+1}. ${s.location} (halt ${s.halt} min)\n`; });
    }
    msg += `Drop: ${data.drop}\n`;
    if (data.tripType === 'roundtrip') {
        let retDate = new Date(data.returnDate).toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
        msg += `Return: ${retDate} at ${data.returnTime}\nReturn Drop: ${data.returnDrop || 'Same as pickup'}\n`;
    }
    if (data.landmark) msg += `Landmark: ${data.landmark}\n`;
    msg += `Date: ${formattedDate}\nTime: ${data.pickupTime}\nPassengers: ${data.passengers}\nLuggage: ${data.luggage}\n\n`;
    msg += `Service Highlights: Free flight tracking, Professional chauffeur, Sanitized EECO Van\n\n`;
    msg += `To confirm, please reply with "OK".`;
    return msg;
}

async function sendEmailCopy(data) {
    try {
        await fetch(`${GAS_URL}?action=sendEmailCopy`, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
        showToast('Confirmation email sent!', 'success');
    } catch (e) { console.log(e); }
}

// ==================== UI HELPERS ====================
function showToast(msg, type = 'success') {
    let toast = document.getElementById('toast');
    if (!toast) { toast = document.createElement('div'); toast.id = 'toast'; toast.className = 'toast'; document.body.appendChild(toast); }
    toast.textContent = msg;
    toast.className = `toast ${type} show`;
    setTimeout(() => toast.classList.remove('show'), 3000);
}

// Global functions for modal buttons
window.closeNoVehiclesModal = closeNoVehiclesModal;
window.closePostSubmitModal = closePostSubmitModal;
window.sendWhatsAppFromModal = sendWhatsAppFromModal;