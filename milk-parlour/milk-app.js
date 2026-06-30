// ==================== MANISH MILK PARLOUR — Shared Config ====================
// After deploying your GAS Web App, paste the URL below:
const GAS_URL = 'YOUR_GAS_WEB_APP_URL_HERE';

// ==================== API HELPER ====================
async function apiGet(params) {
  const url = GAS_URL + '?' + new URLSearchParams(params).toString();
  const res = await fetch(url);
  return res.json();
}

async function apiPost(action, data, token) {
  const url = GAS_URL + '?action=' + action + (token ? '&token=' + token : '');
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action, token, ...data })
  });
  return res.json();
}

// ==================== SESSION HELPERS ====================
function saveSession(key, data, ttlHours = 8) {
  const expiry = Date.now() + ttlHours * 60 * 60 * 1000;
  localStorage.setItem(key, JSON.stringify({ ...data, expiry }));
}

function getSession(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const data = JSON.parse(raw);
    if (data.expiry < Date.now()) { localStorage.removeItem(key); return null; }
    return data;
  } catch (_) { return null; }
}

function clearSession(key) { localStorage.removeItem(key); }

// ==================== TOAST NOTIFICATIONS ====================
function showToast(message, type = 'info', duration = 3500) {
  const existing = document.querySelector('.mp-toast');
  if (existing) existing.remove();
  const t = document.createElement('div');
  t.className = 'mp-toast mp-toast-' + type;
  t.textContent = message;
  document.body.appendChild(t);
  setTimeout(() => t.classList.add('show'), 10);
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 350); }, duration);
}

// ==================== LOADING STATE ====================
function setLoading(btn, loading, text = 'Submit') {
  if (loading) { btn.disabled = true; btn.dataset.origText = btn.textContent; btn.innerHTML = '<span class="spinner"></span> Loading…'; }
  else { btn.disabled = false; btn.textContent = btn.dataset.origText || text; }
}

// ==================== MONTH NAMES ====================
const MONTHS = ['', 'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'];
const MONTHS_SHORT = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
