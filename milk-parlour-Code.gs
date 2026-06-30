// ==================== MANISH MILK PARLOUR — GAS Backend v2 ====================
// Deploy as Web App: Execute as Me, Anyone can access
// Run setupMilkSheets() ONCE after pasting, then deploy.

// ==================== ROUTING ====================
function doGet(e) {
  const a = e?.parameter?.action;
  const t = e?.parameter?.token;
  if (a === 'getConfig')             return getPublicConfig();
  if (a === 'customerLogin')         return customerLogin(e.parameter.phone, e.parameter.pin);
  if (a === 'getCustomerDashboard')  return getCustomerDashboard(t, e.parameter.month, e.parameter.year);
  if (a === 'getCustomerBills')      return getCustomerBills(t);
  if (a === 'deliveryLogin')         return deliveryLogin(e.parameter.personId, e.parameter.password);
  if (a === 'getRoute')              return getRoute(t, e.parameter.date);
  if (a === 'adminLogin')            return adminLogin(e.parameter.password);
  if (a === 'getAdminDashboard')     return getAdminDashboard(t);
  if (a === 'getCustomers')          return getCustomers(t);
  if (a === 'getDeliveryPersons')    return getDeliveryPersons(t);
  if (a === 'getPrices')             return getPrices(t);
  if (a === 'getBills')              return getBills(t, e.parameter.month, e.parameter.year);
  if (a === 'getDeliveries')         return getDeliveries(t, e.parameter.customerId, e.parameter.month, e.parameter.year);
  if (a === 'getFullConfig')         return getFullConfig(t);
  if (a === 'getRegistrations')      return getRegistrations(t);
  if (a === 'getPayments')           return getPayments(t, e.parameter.customerId, e.parameter.month, e.parameter.year);
  if (a === 'getTomorrowOrder')      return getTomorrowOrder(t);
  if (a === 'getEODSummary')         return getEODSummary(t);
  if (a === 'getWholesalePrices')    return getWholesalePrices(t);
  return json({ success: false, error: 'Invalid action' });
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch(_) {}
  const a = e?.parameter?.action || d.action;
  const t = e?.parameter?.token  || d.token;
  if (a === 'markDelivered')          return markDelivered(t, d);
  if (a === 'addCustomer')            return addCustomer(t, d);
  if (a === 'updateCustomer')         return updateCustomer(t, d);
  if (a === 'deleteCustomer')         return deleteCustomer(t, d.customerId);
  if (a === 'addDeliveryPerson')      return addDeliveryPerson(t, d);
  if (a === 'updateDeliveryPerson')   return updateDeliveryPerson(t, d);
  if (a === 'deleteDeliveryPerson')   return deleteDeliveryPerson(t, d.personId);
  if (a === 'updatePrice')            return updatePrice(t, d);
  if (a === 'generateBills')          return generateBills(t, d.month, d.year);
  if (a === 'sendBills')              return sendBills(t, d.month, d.year, d.customerIds);
  if (a === 'sendCredentials')        return sendCredentials(t, d.customerId);
  if (a === 'resetCustomerPin')       return resetCustomerPin(t, d.customerId);
  if (a === 'updateConfig')           return updateConfig(t, d.settings);
  if (a === 'changeAdminPassword')    return changeAdminPassword(t, d.currentPassword, d.newPassword);
  if (a === 'changeDeliveryPassword') return changeDeliveryPassword(t, d.currentPassword, d.newPassword);
  if (a === 'recordPayment')          return recordPayment(t, d);
  if (a === 'submitRegistration')     return submitRegistration(d);
  if (a === 'approveRegistration')    return approveRegistration(t, d.regId);
  if (a === 'rejectRegistration')     return rejectRegistration(t, d.regId, d.reason);
  if (a === 'placeTomorrowOrder')     return placeTomorrowOrder(t, d);
  if (a === 'sendEODReport')          return sendEODReport(t);
  if (a === 'updateWholesalePrice')   return updateWholesalePrice(t, d);
  return json({ success: false, error: 'Invalid action' });
}

// ==================== SHEET SETUP (Run once) ====================
function setupMilkSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');

  const C = { // Color palette
    header: '#1a6b8a', headerTxt: '#FFFFFF',
    altRow: '#EDF6FA',  border: '#BFD8E3',
    active: '#D4EDDA',  inactive: '#F8D7DA',
    pending: '#FFF3CD', delivered: '#D4EDDA', missed: '#F8D7DA',
    accent: '#E8963A',  accentTxt: '#FFFFFF',
    tabColor: '#1a6b8a'
  };

  function styleHeader(sheet, ncols) {
    const hdr = sheet.getRange(1, 1, 1, ncols);
    hdr.setBackground(C.header).setFontColor(C.headerTxt)
       .setFontWeight('bold').setFontSize(11)
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
       .setWrap(true);
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);
    sheet.setTabColor(C.tabColor);
    // Protect header row from accidental edit (warning mode)
    const prot = sheet.getRange(1, 1, 1, ncols).protect();
    prot.setDescription('🔒 Header — Do Not Edit');
    prot.setWarningOnly(true);
  }

  function addRowProtection(sheet) {
    // Protect the sheet from deletion with warning
    const sp = sheet.protect().setDescription('🔒 Sheet protected — use Trash to recover deleted data');
    sp.setWarningOnly(true);
  }

  function dropdownRule(values) {
    return SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(false)
      .setHelpText('Please select a value from the list.')
      .build();
  }

  function altRowFormat(sheet, nrows, ncols) {
    // Apply alternating row colors for readability
    for (let r = 2; r <= Math.max(nrows + 1, 50); r += 2) {
      sheet.getRange(r, 1, 1, ncols).setBackground(C.altRow);
    }
  }

  // ── CONFIG ─────────────────────────────────────────────────────
  let cfg = ss.getSheetByName('Config');
  if (!cfg) {
    cfg = ss.insertSheet('Config');
    cfg.setTabColor('#f39c12');
    const hdr = cfg.getRange('A1:C1');
    hdr.setValues([['⚙️ Setting Key', '📝 Value', '💬 Description']]);
    hdr.setBackground('#f39c12').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    cfg.setFrozenRows(1);
    const prot = cfg.getRange('A1:C1').protect(); prot.setDescription('Header'); prot.setWarningOnly(true);
    const configData = [
      ['BUSINESS_NAME',              'Manish Milk Parlour',     'Business display name'],
      ['BUSINESS_CITY',              'Hubballi',                'City name'],
      ['BUSINESS_PHONE',             '9480324895',              '10-digit WhatsApp/call number'],
      ['BUSINESS_ADDRESS',           'Hubballi, Karnataka',     'Full address for emails/bills'],
      ['UPI_ID',                     '9480324895',              'UPI ID for payment (e.g. name@upi)'],
      ['ADMIN_EMAIL',                '',                        'Admin notification email'],
      ['SUPPORT_EMAIL',              '',                        'Customer support email'],
      ['DRY_RUN_MODE',               'TRUE',                    'TRUE = test mode; all emails go to DRY_RUN_EMAIL'],
      ['DRY_RUN_EMAIL',              '',                        'Your own email for dry-run testing'],
      ['WHATSAPP_PROVIDER',          'twilio',                  'twilio | wati | none'],
      ['WHATSAPP_ACCOUNT_SID',       '',                        'Twilio Account SID'],
      ['WHATSAPP_AUTH_TOKEN',        '',                        'Twilio Auth Token'],
      ['WHATSAPP_FROM',              'whatsapp:+14155238886',   'Twilio sandbox WhatsApp number'],
      ['DELIVERY_CHARGE_PER_MONTH',  '0',                       'Fixed delivery charge per customer per month (₹)'],
      ['GITHUB_PAGES_URL',           '',                        'Full URL of GitHub Pages portal (with trailing /)'],
      ['PORTAL_ACTIVE',              'TRUE',                    'TRUE = portal accessible; FALSE = maintenance mode'],
      ['CURRENCY_SYMBOL',            '₹',                       'Currency symbol for bills'],
      ['TIMEZONE',                   'Asia/Kolkata',            'Timezone for timestamps'],
      ['INVOICE_PREFIX',             'MMP',                     'Prefix for invoice numbers e.g. MMP-2024-0001'],
      ['ORDER_CUTOFF_TIME',          '19:00',                   'Daily cutoff for customers to modify tomorrow\'s order (HH:MM, 24hr)'],
      ['WHOLESALER_PHONE',           '',                        'Wholesaler WhatsApp number (10-digit) for EOD order report'],
      ['WHOLESALER_NAME',            '',                        'Wholesaler / supplier name for report header'],
    ];
    cfg.getRange(2, 1, configData.length, 3).setValues(configData);
    cfg.setColumnWidth(1, 240); cfg.setColumnWidth(2, 280); cfg.setColumnWidth(3, 350);
    cfg.getRange('A2:A20').setFontWeight('bold').setFontColor('#1a6b8a');
    cfg.getRange('C2:C20').setFontColor('#7f8c8d').setFontStyle('italic');
  }

  // ── ADMIN ──────────────────────────────────────────────────────
  if (!ss.getSheetByName('Admin')) {
    const s = ss.insertSheet('Admin');
    s.setTabColor('#e74c3c');
    s.getRange('A1:C1').setValues([['🔑 Field', '🔐 Value (Hashed)', '📅 Last Updated']]);
    s.getRange('A1:C1').setBackground('#e74c3c').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
    s.getRange('A2:C2').setValues([['password_hash', sha256_('milkadmin123'), new Date().toISOString()]]);
    s.getRange('A2').setFontWeight('bold');
    s.setColumnWidths(1, 3, 220);
    s.setFrozenRows(1);
    s.getRange('A2:A10').setNote('Do not edit directly. Use the portal Settings tab to change password.');
    const prot = s.protect().setDescription('🔒 Admin sheet — do not edit manually');
    prot.setWarningOnly(true);
  }

  // ── CUSTOMERS ──────────────────────────────────────────────────
  if (!ss.getSheetByName('Customers')) {
    const s = ss.insertSheet('Customers');
    const h = ['CustomerID','Name','Phone','Email','Address','Shift','DefaultBrand','DefaultQuality','MilkQty/Day','CurdQty/Day','PIN_Hash','Active','Created','LastPINSent','Notes'];
    const widths = [120,180,100,200,220,100,130,130,80,80,300,60,160,160,200];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    // Dropdowns
    s.getRange('F2:F1000').setDataValidation(dropdownRule(['Morning','Evening','Morning+Evening']));
    s.getRange('G2:G1000').setDataValidation(dropdownRule(['Amul','Nandini','Kolhapur','Mother Dairy','Aditya','Amul+Aditya','Kolhapur+Amul','Nandini+Taza']));
    s.getRange('H2:H1000').setDataValidation(dropdownRule(['Gold','Taza','Slim & Trim','Pasteurised Toned','Homogenised','Regular','Full Cream','Gold+Taza','Nandini+Gold','Nandini+Taza']));
    s.getRange('L2:L1000').setDataValidation(dropdownRule(['TRUE','FALSE']));
    // Number validation
    s.getRange('I2:J1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(0, 20).setHelpText('Enter quantity in litres (0–20)').build()
    );
    // Conditional formatting: Active column
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('TRUE') .setBackground(C.active)  .setRanges([s.getRange('L:L')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FALSE').setBackground(C.inactive).setRanges([s.getRange('L:L')]).build(),
    ]);
    altRowFormat(s, 50, h.length);
    addRowProtection(s);
    s.getRange('A1').setNote('CustomerID is auto-generated. Do not edit.');
    s.getRange('K1').setNote('PIN is stored as SHA-256 hash. Never edit directly.');
  }

  // ── DELIVERY PERSONS ───────────────────────────────────────────
  if (!ss.getSheetByName('DeliveryPersons')) {
    const s = ss.insertSheet('DeliveryPersons');
    const h = ['PersonID','Name','Phone','Email','Password_Hash','Active','Created','Route_Sequence'];
    const widths = [100,180,100,200,300,60,160,200];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    s.getRange('F2:F1000').setDataValidation(dropdownRule(['TRUE','FALSE']));
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('TRUE') .setBackground(C.active)  .setRanges([s.getRange('F:F')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FALSE').setBackground(C.inactive).setRanges([s.getRange('F:F')]).build(),
    ]);
    altRowFormat(s, 20, h.length);
    addRowProtection(s);
    s.getRange('E1').setNote('Password stored as SHA-256 hash. Use Admin portal to reset passwords.');
  }

  // ── PRICES ─────────────────────────────────────────────────────
  if (!ss.getSheetByName('Prices')) {
    const s = ss.insertSheet('Prices');
    const h = ['Brand','Quality / Variant','Price per Litre (₹)','Last Updated'];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    s.setColumnWidths(1, 4, 180);
    s.getRange('A2:A1000').setDataValidation(dropdownRule(['Amul','Nandini','Kolhapur','Mother Dairy','Aditya','Amul+Aditya','Kolhapur+Amul','Nandini+Taza']));
    s.getRange('C2:C1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(1, 500).setHelpText('Price in ₹ per litre').build()
    );
    s.getRange('C2:C1000').setNumberFormat('₹#,##0.00');
    const defaults = [
      ['Amul','Gold',68,''],['Amul','Taza',54,''],['Amul','Slim & Trim',52,''],
      ['Nandini','Pasteurised Toned',44,''],['Nandini','Homogenised',50,''],
      ['Kolhapur','Regular',52,''],['Mother Dairy','Full Cream',64,''],
      ['Mother Dairy','Toned',52,''],['Aditya','Regular',50,''],
    ];
    s.getRange(2,1,defaults.length,4).setValues(defaults);
    altRowFormat(s, defaults.length + 10, h.length);
    addRowProtection(s);
    s.setFrozenRows(1);
  }

  // ── DELIVERIES ─────────────────────────────────────────────────
  if (!ss.getSheetByName('Deliveries')) {
    const s = ss.insertSheet('Deliveries');
    const h = ['DeliveryID','CustomerID','CustomerName','Date','Time','Day','Month','Year','PartnerID','PartnerName','Brand','Quality','MilkQty','CurdQty','Status','Notes'];
    const widths = [140,120,180,100,80,50,60,60,100,180,130,130,70,70,90,200];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    s.getRange('O2:O1000').setDataValidation(dropdownRule(['Delivered','Missed']));
    s.getRange('M2:N1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(0,20).setHelpText('Quantity in litres').build()
    );
    s.getRange('M2:N1000').setNumberFormat('0.00');
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Delivered').setBackground(C.delivered).setRanges([s.getRange('O:O')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Missed')   .setBackground(C.missed)   .setRanges([s.getRange('O:O')]).build(),
    ]);
    s.setFrozenRows(1);
    s.setFrozenColumns(3);
    addRowProtection(s);
  }

  // ── BILLS ──────────────────────────────────────────────────────
  if (!ss.getSheetByName('Bills')) {
    const s = ss.insertSheet('Bills');
    const h = ['BillID','CustomerID','CustomerName','Phone','Email','Month','Year','TotalDays','MilkLtrs','CurdLtrs','MilkAmount','CurdAmount','DeliveryCharge','TotalAmount','GeneratedAt','EmailSentAt','WhatsAppSentAt','Status'];
    const widths = [180,120,180,100,200,60,60,70,80,80,100,100,100,120,160,160,160,100];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    s.getRange('N2:N1000').setNumberFormat('₹#,##0.00');
    s.getRange('K2:M1000').setNumberFormat('₹#,##0.00');
    s.getRange('I2:J1000').setNumberFormat('0.00');
    s.getRange('R2:R1000').setDataValidation(dropdownRule(['Generated','Sent','Partially Sent']));
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Sent')            .setBackground(C.active)  .setRanges([s.getRange('R:R')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Generated')       .setBackground(C.pending) .setRanges([s.getRange('R:R')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Partially Sent')  .setBackground(C.altRow)  .setRanges([s.getRange('R:R')]).build(),
    ]);
    s.setFrozenRows(1);
    s.setFrozenColumns(3);
    addRowProtection(s);
  }

  // ── PAYMENTS ───────────────────────────────────────────────────
  if (!ss.getSheetByName('Payments')) {
    const s = ss.insertSheet('Payments');
    const h = ['PaymentID','BillID','CustomerID','CustomerName','Phone','Email','Month','Year','Amount','PaymentMethod','TransactionRef','ReceivedAt','ReceiptEmailAt','ReceiptWhatsAppAt','RecordedBy'];
    const widths = [160,180,120,180,100,200,60,60,120,130,180,160,160,160,150];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    s.getRange('I2:I1000').setNumberFormat('₹#,##0.00');
    s.getRange('J2:J1000').setDataValidation(dropdownRule(['Cash','UPI / GPay','PhonePe','Paytm','NEFT / IMPS','Cheque','Other']));
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Cash').setBackground('#e8f5e9').setRanges([s.getRange('J:J')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('UPI / GPay').setBackground('#e3f2fd').setRanges([s.getRange('J:J')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PhonePe').setBackground('#f3e5f5').setRanges([s.getRange('J:J')]).build(),
    ]);
    s.setFrozenRows(1);
    s.setFrozenColumns(4);
    addRowProtection(s);
  }

  // ── REGISTRATIONS ──────────────────────────────────────────────
  if (!ss.getSheetByName('Registrations')) {
    const s = ss.insertSheet('Registrations');
    const h = ['RegID','Name','Phone','Email','Address','PreferredBrand','PreferredQty','Shift','Message','RegisteredAt','Status','ProcessedAt','AdminNotes'];
    const widths = [140,180,100,200,220,130,80,100,250,160,100,160,250];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    s.getRange('K2:K1000').setDataValidation(dropdownRule(['Pending','Approved','Rejected']));
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending') .setBackground(C.pending) .setRanges([s.getRange('K:K')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Approved').setBackground(C.active)  .setRanges([s.getRange('K:K')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Rejected').setBackground(C.inactive).setRanges([s.getRange('K:K')]).build(),
    ]);
    s.setFrozenRows(1);
    addRowProtection(s);
  }

  // ── TOMORROW ORDERS ────────────────────────────────────────────
  if (!ss.getSheetByName('TomorrowOrders')) {
    const s = ss.insertSheet('TomorrowOrders');
    const h = ['OrderID','CustomerID','CustomerName','Phone','ForDate','Brand','Quality','MilkQty','CurdQty','PlacedAt','ModifiedAt'];
    const widths = [160,120,180,100,100,130,130,70,70,160,160];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    h.forEach((_, i) => s.setColumnWidth(i+1, widths[i]));
    s.getRange('H2:I1000').setNumberFormat('0.00');
    s.getRange('H2:I1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(0,20).setHelpText('Litres').build()
    );
    s.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenDateEqualTo(SpreadsheetApp.RelativeDate.TODAY)
        .setBackground('#e8f5e9').setRanges([s.getRange('E:E')]).build(),
    ]);
    s.setFrozenRows(1);
    s.setFrozenColumns(3);
    addRowProtection(s);
    s.getRange('A1').setNote('One row per customer per date. Customer can modify before cutoff (7 PM). Delivery person sees this as the requested order.');
  }

  // ── WHOLESALE PRICES ───────────────────────────────────────────
  if (!ss.getSheetByName('WholesalePrices')) {
    const s = ss.insertSheet('WholesalePrices');
    const h = ['Brand','Quality / Variant','Wholesale Price/Ltr (₹)','Last Updated'];
    s.getRange(1,1,1,h.length).setValues([h]);
    styleHeader(s, h.length);
    s.setColumnWidths(1, 4, 180);
    s.getRange('A2:A1000').setDataValidation(dropdownRule(['Amul','Nandini','Kolhapur','Mother Dairy','Aditya','Amul+Aditya','Kolhapur+Amul','Nandini+Taza']));
    s.getRange('C2:C1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(1,500).setHelpText('Wholesale price ₹/litre').build()
    );
    s.getRange('C2:C1000').setNumberFormat('₹#,##0.00');
    // Seed with typical wholesale prices (roughly 10-15% below retail)
    const wDefaults = [
      ['Amul','Gold',60,''],['Amul','Taza',47,''],['Amul','Slim & Trim',45,''],
      ['Nandini','Pasteurised Toned',38,''],['Nandini','Homogenised',44,''],
      ['Kolhapur','Regular',45,''],['Mother Dairy','Full Cream',56,''],
      ['Mother Dairy','Toned',45,''],['Aditya','Regular',43,''],
    ];
    s.getRange(2,1,wDefaults.length,4).setValues(wDefaults);
    altRowFormat(s, wDefaults.length + 10, h.length);
    s.setFrozenRows(1);
    addRowProtection(s);
    s.getRange('C1').setNote('Enter what you pay the wholesaler per litre. Used in the EOD order report to calculate your purchase cost.');
  }

  // Re-order sheets for usability
  const order = ['Config','Admin','Customers','DeliveryPersons','Prices','WholesalePrices','Deliveries','TomorrowOrders','Bills','Payments','Registrations'];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh), ss.moveActiveSheet(i + 1);
  });

  SpreadsheetApp.getUi().alert(
    '✅ Manish Milk Parlour — Sheet Setup Complete!\n\n' +
    'Next steps:\n' +
    '1. Fill Config sheet: ADMIN_EMAIL, DRY_RUN_EMAIL, BUSINESS_PHONE, UPI_ID\n' +
    '2. Deploy as Web App (Execute as Me, Anyone access)\n' +
    '3. Copy Web App URL → paste into milk-app.js\n\n' +
    'Default admin password: milkadmin123 (change via portal Settings)\n\n' +
    '⚠️ IMPORTANT: All sheets are warning-protected.\n' +
    'Deleted data can be recovered: File → Version History → See version history'
  );
}

// ==================== UTILITIES ====================
function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function sha256_(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function getConfigMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) return {};
  const data = sheet.getRange('A:B').getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) { if (data[i][0]) cfg[data[i][0]] = data[i][1]; }
  return cfg;
}
function cfg_(key) { return getConfigMap_()[key] || ''; }
function isDryRun_() { return cfg_('DRY_RUN_MODE') === 'TRUE'; }
function resolveEmail_(email) { return isDryRun_() ? cfg_('DRY_RUN_EMAIL') : email; }
function resolvePhone_(phone)  { return isDryRun_() ? cfg_('BUSINESS_PHONE') : phone; }
function nowIST_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'); }
function todayIST_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'); }
function monthYearIST_() {
  return { month: parseInt(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'M')),
           year:  parseInt(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy')) };
}
function monthName_(m) { return ['','January','February','March','April','May','June','July','August','September','October','November','December'][parseInt(m)]; }
function generatePin_() { return String(Math.floor(100000 + Math.random() * 900000)); }
function nextInvoiceId_() {
  const props = PropertiesService.getScriptProperties();
  let n = parseInt(props.getProperty('invoice_seq') || '0') + 1;
  props.setProperty('invoice_seq', n.toString());
  const prefix = cfg_('INVOICE_PREFIX') || 'MMP';
  return `${prefix}-${new Date().getFullYear()}-${String(n).padStart(4,'0')}`;
}

// ==================== TOKEN AUTH ====================
function issueToken_(scope, id) {
  const props = PropertiesService.getScriptProperties();
  const token = Utilities.getUuid();
  const expiry = Date.now() + 8 * 60 * 60 * 1000;
  props.setProperty(`token_${token}`, JSON.stringify({ scope, id, expiry }));
  cleanTokens_();
  return token;
}
function validateToken_(token, requiredScope) {
  if (!token) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(`token_${token}`);
  if (!raw) return null;
  try {
    const data = JSON.parse(raw);
    if (data.expiry < Date.now()) { props.deleteProperty(`token_${token}`); return null; }
    if (requiredScope && data.scope !== requiredScope && data.scope !== 'admin') return null;
    return data;
  } catch(_) { return null; }
}
function cleanTokens_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  for (const k in all) {
    if (k.startsWith('token_')) {
      try { const d = JSON.parse(all[k]); if (d.expiry < now) props.deleteProperty(k); }
      catch(_) { props.deleteProperty(k); }
    }
  }
}

// ==================== AUTH ENDPOINTS ====================
function adminLogin(password) {
  const adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin');
  if (!adminSheet) return json({ success: false, error: 'Admin not configured. Run setupMilkSheets first.' });
  if (sha256_(password) !== adminSheet.getRange('B2').getValue()) return json({ success: false, error: 'Invalid password' });
  return json({ success: true, token: issueToken_('admin','admin'), businessName: cfg_('BUSINESS_NAME') });
}
function changeAdminPassword(token, cur, nw) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin');
  if (sha256_(cur) !== s.getRange('B2').getValue()) return json({ success:false, error:'Current password incorrect' });
  s.getRange('B2').setValue(sha256_(nw)); s.getRange('C2').setValue(new Date().toISOString());
  return json({ success: true });
}
function customerLogin(phone, pin) {
  if (!phone || !pin) return json({ success:false, error:'Phone and PIN required' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers')?.getDataRange().getValues() || [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(phone) && data[i][11] === 'TRUE' && sha256_(String(pin)) === String(data[i][10])) {
      return json({ success:true, token: issueToken_('customer', data[i][0]), customerId: data[i][0], name: data[i][1] });
    }
  }
  return json({ success:false, error:'Invalid phone or PIN' });
}
function deliveryLogin(personId, password) {
  if (!personId || !password) return json({ success:false, error:'ID and password required' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeliveryPersons')?.getDataRange().getValues() || [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(personId) && data[i][5] === 'TRUE' && sha256_(String(password)) === String(data[i][4])) {
      return json({ success:true, token: issueToken_('delivery', data[i][0]), personId: data[i][0], name: data[i][1] });
    }
  }
  return json({ success:false, error:'Invalid ID or password' });
}
function changeDeliveryPassword(token, cur, nw) {
  const sess = validateToken_(token,'delivery');
  if (!sess) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeliveryPersons');
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sess.id)) {
      if (sha256_(cur) !== String(data[i][4])) return json({ success:false, error:'Current password incorrect' });
      s.getRange(i+1,5).setValue(sha256_(nw));
      return json({ success:true });
    }
  }
  return json({ success:false, error:'Not found' });
}

// ==================== CONFIG ====================
function getPublicConfig() {
  const c = getConfigMap_();
  return json({ success:true, businessName:c.BUSINESS_NAME||'Manish Milk Parlour', businessCity:c.BUSINESS_CITY||'Hubballi',
    businessPhone:c.BUSINESS_PHONE||'', portalActive:c.PORTAL_ACTIVE!=='FALSE', dryRunMode:c.DRY_RUN_MODE==='TRUE',
    currencySymbol:c.CURRENCY_SYMBOL||'₹' });
}
function getFullConfig(token) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const c = getConfigMap_(); delete c.WHATSAPP_ACCOUNT_SID; delete c.WHATSAPP_AUTH_TOKEN;
  return json({ success:true, config:c });
}
function updateConfig(token, settings) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = sheet.getRange('A:B').getValues();
  for (const [key, value] of Object.entries(settings)) {
    let found = false;
    for (let i = 1; i < data.length; i++) { if (data[i][0] === key) { sheet.getRange(i+1,2).setValue(value); found=true; break; } }
    if (!found) sheet.appendRow([key, value, 'Custom setting']);
  }
  return json({ success:true });
}

// ==================== CUSTOMERS ====================
function getCustomers(token) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers')?.getDataRange().getValues() || [];
  const headers = data[0] || [];
  return json({ success:true, customers: data.slice(1).filter(r=>r[0]).map(r => {
    const row = {}; headers.forEach((h,j) => row[h]=r[j]); delete row['PIN_Hash']; return row;
  })});
}
function addCustomer(token, d) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const id = 'CUST' + Utilities.formatDate(new Date(),'Asia/Kolkata','yyMMddHHmmss');
  const pin = generatePin_();
  s.appendRow([id, d.name, d.phone, d.email||'', d.address||'', d.shift||'Morning',
    d.defaultBrand||'Amul', d.defaultQuality||'Gold', d.defaultMilkQty||1, d.defaultCurdQty||0,
    sha256_(pin), 'TRUE', nowIST_(), '', d.notes||'']);
  const emailSent = d.email ? sendCredentialEmail_(id, d.name, d.phone, pin, d.email) : false;
  return json({ success:true, customerId:id, pin, pinSent:emailSent });
}
function updateCustomer(token, d) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(d.customerId)) {
      const map = {name:2,phone:3,email:4,address:5,shift:6,defaultBrand:7,defaultQuality:8,defaultMilkQty:9,defaultCurdQty:10,active:12,notes:15};
      for (const [key,col] of Object.entries(map)) {
        if (d[key]!==undefined) s.getRange(i+1,col).setValue(key==='active'?(d[key]?'TRUE':'FALSE'):d[key]);
      }
      return json({ success:true });
    }
  }
  return json({ success:false, error:'Customer not found' });
}
function deleteCustomer(token, customerId) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(customerId)) { s.getRange(i+1,12).setValue('FALSE'); return json({ success:true }); }
  }
  return json({ success:false, error:'Not found' });
}
function resetCustomerPin(token, customerId) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(customerId)) {
      const pin = generatePin_();
      s.getRange(i+1,11).setValue(sha256_(pin));
      s.getRange(i+1,14).setValue(nowIST_());
      const emailSent = data[i][3] ? sendCredentialEmail_(customerId, data[i][1], data[i][2], pin, data[i][3]) : false;
      return json({ success:true, pin, pinSent:emailSent });
    }
  }
  return json({ success:false, error:'Not found' });
}
function sendCredentials(token, customerId) { return resetCustomerPin(token, customerId); }

// ==================== DELIVERY PERSONS ====================
function getDeliveryPersons(token) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeliveryPersons')?.getDataRange().getValues()||[];
  const headers = data[0]||[];
  return json({ success:true, persons: data.slice(1).filter(r=>r[0]).map(r => {
    const row={}; headers.forEach((h,j)=>row[h]=r[j]); delete row['Password_Hash']; return row;
  })});
}
function addDeliveryPerson(token, d) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeliveryPersons');
  const existing = s.getDataRange().getValues();
  const id = 'DP' + String(existing.length).padStart(3,'0');
  const password = d.password || String(Math.floor(10000+Math.random()*90000));
  s.appendRow([id, d.name, d.phone, d.email||'', sha256_(password), 'TRUE', nowIST_(), d.sequence||'']);
  return json({ success:true, personId:id, password });
}
function updateDeliveryPerson(token, d) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeliveryPersons');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(d.personId)) {
      if (d.name!==undefined)     s.getRange(i+1,2).setValue(d.name);
      if (d.phone!==undefined)    s.getRange(i+1,3).setValue(d.phone);
      if (d.email!==undefined)    s.getRange(i+1,4).setValue(d.email);
      if (d.password!==undefined) s.getRange(i+1,5).setValue(sha256_(d.password));
      if (d.active!==undefined)   s.getRange(i+1,6).setValue(d.active?'TRUE':'FALSE');
      if (d.sequence!==undefined) s.getRange(i+1,8).setValue(d.sequence);
      return json({ success:true });
    }
  }
  return json({ success:false, error:'Not found' });
}
function deleteDeliveryPerson(token, personId) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DeliveryPersons');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0])===String(personId)) { s.getRange(i+1,6).setValue('FALSE'); return json({ success:true }); }
  }
  return json({ success:false, error:'Not found' });
}

// ==================== PRICES ====================
function getPrices(token) {
  const sess = validateToken_(token,'admin')||validateToken_(token,'delivery')||validateToken_(token,'customer');
  if (!sess) return json({ success:false, error:'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prices')?.getDataRange().getValues()||[];
  return json({ success:true, prices: data.slice(1).filter(r=>r[0]).map(r=>({ brand:r[0], quality:r[1], pricePerLitre:r[2] })) });
}
function updatePrice(token, d) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Prices');
  const data = s.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0]===d.brand && data[i][1]===d.quality) {
      s.getRange(i+1,3).setValue(d.pricePerLitre); s.getRange(i+1,4).setValue(nowIST_()); return json({ success:true });
    }
  }
  s.appendRow([d.brand, d.quality, d.pricePerLitre, nowIST_()]);
  return json({ success:true });
}

// ==================== DELIVERY ROUTE ====================
function getRoute(token, date) {
  const sess = validateToken_(token,'delivery')||validateToken_(token,'admin');
  if (!sess) return json({ success:false, error:'Unauthorised' });
  const targetDate = date||todayIST_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const custData   = ss.getSheetByName('Customers')?.getDataRange().getValues()||[];
  const delivData  = ss.getSheetByName('Deliveries')?.getDataRange().getValues()||[];
  const markedToday = {};
  for (let i=1; i<delivData.length; i++) {
    if (String(delivData[i][3])===String(targetDate)) {
      markedToday[String(delivData[i][1])] = { status:delivData[i][14], time:delivData[i][4] };
    }
  }
  // Check TomorrowOrders placed for today (orders placed yesterday evening)
  const orderData = ss.getSheetByName('TomorrowOrders')?.getDataRange().getValues()||[];
  const requestedOrders = {};
  for (let i=1; i<orderData.length; i++) {
    if (String(orderData[i][4])===String(targetDate)) {
      requestedOrders[String(orderData[i][1])] = {
        brand:orderData[i][5], quality:orderData[i][6],
        milkQty:parseFloat(orderData[i][7])||0, curdQty:parseFloat(orderData[i][8])||0
      };
    }
  }
  const route = custData.slice(1).filter(r=>r[11]==='TRUE').map(r=>({
    customerId:String(r[0]), name:r[1], phone:r[2], address:r[4], shift:r[5],
    defaultBrand:r[6], defaultQuality:r[7], defaultMilkQty:r[8], defaultCurdQty:r[9],
    requestedOrder:requestedOrders[String(r[0])]||null,
    todayStatus:markedToday[String(r[0])]||null
  }));
  return json({ success:true, route, date:targetDate, personId:sess.id });
}

// ==================== MARK DELIVERY ====================
function markDelivered(token, d) {
  const sess = validateToken_(token,'delivery')||validateToken_(token,'admin');
  if (!sess) return json({ success:false, error:'Unauthorised' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Deliveries');
  const custData = ss.getSheetByName('Customers')?.getDataRange().getValues()||[];
  let custName=''; for(let i=1;i<custData.length;i++){if(String(custData[i][0])===String(d.customerId)){custName=custData[i][1];break;}}
  let personName='Admin';
  if (sess.scope==='delivery') {
    const dp = ss.getSheetByName('DeliveryPersons')?.getDataRange().getValues()||[];
    for(let i=1;i<dp.length;i++){if(String(dp[i][0])===String(sess.id)){personName=dp[i][1];break;}}
  }
  const targetDate = d.date||todayIST_();
  const now = nowIST_();
  const [ty,tm,td2] = targetDate.split('-');
  const id = 'DEL'+Utilities.formatDate(new Date(),'Asia/Kolkata','yyMMddHHmmss')+Math.floor(Math.random()*100);
  const existing = sheet.getDataRange().getValues();
  for (let i=1; i<existing.length; i++) {
    if (String(existing[i][1])===String(d.customerId) && String(existing[i][3])===String(targetDate)) {
      sheet.getRange(i+1,5).setValue(now.split(' ')[1]);
      sheet.getRange(i+1,11).setValue(d.brand); sheet.getRange(i+1,12).setValue(d.quality||'');
      sheet.getRange(i+1,13).setValue(d.milkQty||0); sheet.getRange(i+1,14).setValue(d.curdQty||0);
      sheet.getRange(i+1,15).setValue(d.status||'Delivered'); sheet.getRange(i+1,16).setValue(d.notes||'');
      return json({ success:true, deliveryId:existing[i][0], updated:true });
    }
  }
  sheet.appendRow([id, d.customerId, custName, targetDate, now.split(' ')[1],
    parseInt(td2), parseInt(tm), parseInt(ty), sess.id, personName,
    d.brand, d.quality||'', d.milkQty||0, d.curdQty||0, d.status||'Delivered', d.notes||'']);
  return json({ success:true, deliveryId:id, updated:false });
}

// ==================== CUSTOMER DASHBOARD ====================
function getCustomerDashboard(token, month, year) {
  const sess = validateToken_(token,'customer');
  if (!sess) return json({ success:false, error:'Unauthorised' });
  const {month:cm, year:cy} = monthYearIST_();
  const m=parseInt(month)||cm, y=parseInt(year)||cy;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const custData  = ss.getSheetByName('Customers')?.getDataRange().getValues()||[];
  const delivData = ss.getSheetByName('Deliveries')?.getDataRange().getValues()||[];
  const billData  = ss.getSheetByName('Bills')?.getDataRange().getValues()||[];
  let customer=null;
  for(let i=1;i<custData.length;i++){
    if(String(custData[i][0])===String(sess.id)){
      customer={customerId:custData[i][0],name:custData[i][1],phone:custData[i][2],address:custData[i][4],
        shift:custData[i][5],defaultBrand:custData[i][6],defaultQuality:custData[i][7],
        defaultMilkQty:custData[i][8],defaultCurdQty:custData[i][9]}; break;
    }
  }
  if (!customer) return json({ success:false, error:'Customer not found' });
  const calendar={};let totalMilk=0,totalCurd=0,totalDays=0;
  for(let i=1;i<delivData.length;i++){
    if(String(delivData[i][1])===String(sess.id)&&parseInt(delivData[i][6])===m&&parseInt(delivData[i][7])===y){
      const day=parseInt(delivData[i][5]);
      calendar[day]={status:delivData[i][14],time:delivData[i][4],brand:delivData[i][10],milkQty:parseFloat(delivData[i][12])||0,curdQty:parseFloat(delivData[i][13])||0};
      if(delivData[i][14]==='Delivered'){totalMilk+=parseFloat(delivData[i][12])||0;totalCurd+=parseFloat(delivData[i][13])||0;totalDays++;}
    }
  }
  let bill=null;
  for(let i=1;i<billData.length;i++){
    if(String(billData[i][1])===String(sess.id)&&parseInt(billData[i][5])===m&&parseInt(billData[i][6])===y){
      bill={billId:billData[i][0],totalAmount:billData[i][13],generatedAt:billData[i][14],status:billData[i][17]}; break;
    }
  }
  return json({ success:true, customer, month:m, year:y, monthName:monthName_(m),
    daysInMonth:new Date(y,m,0).getDate(), calendar, totalMilk, totalCurd, totalDays, bill });
}
function getCustomerBills(token) {
  const sess = validateToken_(token,'customer');
  if (!sess) return json({ success:false, error:'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bills')?.getDataRange().getValues()||[];
  const bills = data.slice(1).filter(r=>String(r[1])===String(sess.id)).map(r=>({
    billId:r[0],month:r[5],year:r[6],totalDays:r[7],milkLtrs:r[8],curdLtrs:r[9],totalAmount:r[13],status:r[17],generatedAt:r[14]
  }));
  bills.sort((a,b)=>(b.year*100+b.month)-(a.year*100+a.month));
  return json({ success:true, bills });
}

// ==================== DELIVERIES (ADMIN) ====================
function getDeliveries(token, customerId, month, year) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Deliveries')?.getDataRange().getValues()||[];
  const headers = data[0]||[];
  const deliveries = data.slice(1).filter(r=>r[0]&&
    (!customerId||String(r[1])===String(customerId))&&
    (!month||parseInt(r[6])===parseInt(month))&&
    (!year||parseInt(r[7])===parseInt(year))
  ).map(r=>{ const row={}; headers.forEach((h,j)=>row[h]=r[j]); return row; });
  return json({ success:true, deliveries });
}

// ==================== ADMIN DASHBOARD ====================
function getAdminDashboard(token) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const {month:m,year:y} = monthYearIST_();
  const today = todayIST_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const custData  = ss.getSheetByName('Customers')?.getDataRange().getValues()||[];
  const delivData = ss.getSheetByName('Deliveries')?.getDataRange().getValues()||[];
  const billData  = ss.getSheetByName('Bills')?.getDataRange().getValues()||[];
  const regData   = ss.getSheetByName('Registrations')?.getDataRange().getValues()||[];
  let totalCustomers=0,activeCustomers=0;
  for(let i=1;i<custData.length;i++){if(!custData[i][0])continue;totalCustomers++;if(custData[i][11]==='TRUE')activeCustomers++;}
  let todayDeliveries=0,todayMissed=0,monthDeliveries=0,monthMilkLtrs=0;
  for(let i=1;i<delivData.length;i++){
    if(!delivData[i][0])continue;
    if(String(delivData[i][3])===String(today)){if(delivData[i][14]==='Delivered')todayDeliveries++;else if(delivData[i][14]==='Missed')todayMissed++;}
    if(parseInt(delivData[i][6])===m&&parseInt(delivData[i][7])===y&&delivData[i][14]==='Delivered'){monthDeliveries++;monthMilkLtrs+=parseFloat(delivData[i][12])||0;}
  }
  let monthBillTotal=0,pendingBills=0;
  for(let i=1;i<billData.length;i++){
    if(!billData[i][0])continue;
    if(parseInt(billData[i][5])===m&&parseInt(billData[i][6])===y)monthBillTotal+=parseFloat(billData[i][13])||0;
    if(billData[i][17]==='Generated')pendingBills++;
  }
  const pendingRegs = regData.slice(1).filter(r=>r[0]&&r[10]==='Pending').length;
  return json({ success:true,
    stats:{totalCustomers,activeCustomers,todayDeliveries,todayMissed,monthDeliveries,
      monthMilkLtrs:Math.round(monthMilkLtrs*100)/100,monthBillTotal:Math.round(monthBillTotal),pendingBills,pendingRegs},
    currentMonth:m, currentYear:y, today });
}

// ==================== BILL GENERATION ====================
function generateBills(token, month, year) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const {month:cm,year:cy}=monthYearIST_();
  const m=parseInt(month)||cm, y=parseInt(year)||cy;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const custSheet  = ss.getSheetByName('Customers');
  const delivSheet = ss.getSheetByName('Deliveries');
  const priceSheet = ss.getSheetByName('Prices');
  const billSheet  = ss.getSheetByName('Bills');
  const custData   = custSheet?.getDataRange().getValues()||[];
  const delivData  = delivSheet?.getDataRange().getValues()||[];
  const priceData  = priceSheet?.getDataRange().getValues()||[];
  const priceMap={};
  for(let i=1;i<priceData.length;i++){if(priceData[i][0])priceMap[`${priceData[i][0]}|${priceData[i][1]}`]=parseFloat(priceData[i][2])||0;}
  const delivCharge=parseFloat(cfg_('DELIVERY_CHARGE_PER_MONTH'))||0;
  const now=nowIST_();
  const existing=billSheet.getDataRange().getValues();
  for(let i=existing.length-1;i>=1;i--){if(parseInt(existing[i][5])===m&&parseInt(existing[i][6])===y)billSheet.deleteRow(i+1);}
  const generated=[];
  for(let ci=1;ci<custData.length;ci++){
    if(!custData[ci][0]||custData[ci][11]!=='TRUE')continue;
    const custId=String(custData[ci][0]);
    let totalDays=0,milkLtrs=0,curdLtrs=0,milkAmount=0,curdAmount=0;
    for(let di=1;di<delivData.length;di++){
      if(String(delivData[di][1])!==custId||parseInt(delivData[di][6])!==m||parseInt(delivData[di][7])!==y||delivData[di][14]!=='Delivered')continue;
      totalDays++;
      const mQty=parseFloat(delivData[di][12])||0, cQty=parseFloat(delivData[di][13])||0;
      const price=priceMap[`${delivData[di][10]}|${delivData[di][11]}`]||priceMap[`${delivData[di][10]}|`]||0;
      milkLtrs+=mQty; curdLtrs+=cQty; milkAmount+=mQty*price; curdAmount+=cQty*price;
    }
    const totalAmount=Math.round(milkAmount+curdAmount+delivCharge);
    const billId=`BILL-${y}${String(m).padStart(2,'0')}-${custId}`;
    billSheet.appendRow([billId,custId,custData[ci][1],custData[ci][2],custData[ci][3],
      m,y,totalDays,Math.round(milkLtrs*100)/100,Math.round(curdLtrs*100)/100,
      Math.round(milkAmount),Math.round(curdAmount),delivCharge,totalAmount,now,'','','Generated']);
    generated.push({billId,customerId:custId,name:custData[ci][1],totalAmount,totalDays});
  }
  return json({ success:true, generated, count:generated.length, month:m, year:y, monthName:monthName_(m) });
}
function getBills(token, month, year) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const {month:cm,year:cy}=monthYearIST_();
  const m=parseInt(month)||cm, y=parseInt(year)||cy;
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bills')?.getDataRange().getValues()||[];
  const headers=data[0]||[];
  const bills=data.slice(1).filter(r=>r[0]&&parseInt(r[5])===m&&parseInt(r[6])===y).map(r=>{const row={};headers.forEach((h,j)=>row[h]=r[j]);return row;});
  return json({ success:true, bills, month:m, year:y, monthName:monthName_(m) });
}
function sendBills(token, month, year, customerIds) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const {month:cm,year:cy}=monthYearIST_();
  const m=parseInt(month)||cm, y=parseInt(year)||cy;
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const billSheet=ss.getSheetByName('Bills');
  const data=billSheet.getDataRange().getValues();
  const results=[];
  for(let i=1;i<data.length;i++){
    if(!data[i][0]||parseInt(data[i][5])!==m||parseInt(data[i][6])!==y)continue;
    if(customerIds&&customerIds.length&&!customerIds.includes(String(data[i][1])))continue;
    const bill={billId:data[i][0],custId:data[i][1],name:data[i][2],phone:data[i][3],email:data[i][4],
      month:m,year:y,totalDays:data[i][7],milkLtrs:data[i][8],curdLtrs:data[i][9],
      milkAmount:data[i][10],curdAmount:data[i][11],deliveryCharge:data[i][12],totalAmount:data[i][13]};
    const emailSent=sendBillEmail_(bill);
    const waSent=sendBillWhatsApp_(bill);
    const now=nowIST_();
    if(emailSent)billSheet.getRange(i+1,16).setValue(now);
    if(waSent)billSheet.getRange(i+1,17).setValue(now);
    billSheet.getRange(i+1,18).setValue('Sent');
    results.push({name:bill.name,emailSent,waSent});
  }
  return json({ success:true, results, count:results.length });
}

// ==================== PAYMENTS ====================
function recordPayment(token, d) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paySheet  = ss.getSheetByName('Payments');
  const billSheet = ss.getSheetByName('Bills');
  const custSheet = ss.getSheetByName('Customers');
  const custData  = custSheet?.getDataRange().getValues()||[];
  let custName='',custEmail='',custPhone='';
  for(let i=1;i<custData.length;i++){if(String(custData[i][0])===String(d.customerId)){custName=custData[i][1];custEmail=custData[i][3];custPhone=custData[i][2];break;}}
  const now=nowIST_();
  const payId='PAY'+Utilities.formatDate(new Date(),'Asia/Kolkata','yyMMddHHmmss');
  const {month:cm,year:cy}=monthYearIST_();
  paySheet.appendRow([payId, d.billId||'', d.customerId, custName, custPhone, custEmail,
    d.month||cm, d.year||cy, d.amount, d.method, d.transactionRef||'', now, '', '', d.recordedBy||'Admin']);
  // Send receipt
  const payInfo={payId, custName, custEmail, custPhone, billId:d.billId||'',
    month:d.month||cm, year:d.year||cy, amount:d.amount, method:d.method, transactionRef:d.transactionRef||''};
  const emailSent = custEmail ? sendPaymentReceiptEmail_(payInfo) : false;
  const waSent    = custPhone ? sendPaymentReceiptWhatsApp_(payInfo) : false;
  if(emailSent){ const lr=paySheet.getLastRow(); paySheet.getRange(lr,13).setValue(now); }
  if(waSent)  { const lr=paySheet.getLastRow(); paySheet.getRange(lr,14).setValue(now); }
  // Mark bill as paid if billId given
  if(d.billId&&billSheet){
    const billData=billSheet.getDataRange().getValues();
    for(let i=1;i<billData.length;i++){
      if(String(billData[i][0])===String(d.billId)){billSheet.getRange(i+1,18).setValue('Paid');break;}
    }
  }
  return json({ success:true, paymentId:payId, emailSent, waSent });
}
function getPayments(token, customerId, month, year) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments')?.getDataRange().getValues()||[];
  const headers=data[0]||[];
  const payments=data.slice(1).filter(r=>r[0]&&
    (!customerId||String(r[2])===String(customerId))&&
    (!month||parseInt(r[6])===parseInt(month))&&
    (!year||parseInt(r[7])===parseInt(year))
  ).map(r=>{const row={};headers.forEach((h,j)=>row[h]=r[j]);return row;});
  return json({ success:true, payments });
}

// ==================== REGISTRATIONS ====================
function submitRegistration(d) {
  if (!d.name||!d.phone) return json({ success:false, error:'Name and phone are required' });
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const s=ss.getSheetByName('Registrations');
  if (!s) return json({ success:false, error:'System not ready' });
  // Check for duplicate phone
  const existing=s.getDataRange().getValues();
  for(let i=1;i<existing.length;i++){
    if(String(existing[i][2])===String(d.phone)&&existing[i][10]==='Pending')
      return json({ success:false, error:'A request with this phone number is already pending.' });
  }
  const regId='REG'+Utilities.formatDate(new Date(),'Asia/Kolkata','yyMMddHHmmss');
  s.appendRow([regId, d.name, d.phone, d.email||'', d.address||'', d.preferredBrand||'Amul',
    d.preferredQty||1, d.shift||'Morning', d.message||'', nowIST_(), 'Pending','','']);
  // Thank-you to applicant
  if (d.email) sendRegistrationAckEmail_(d.name, d.email);
  // Notify admin
  sendRegistrationAdminAlert_(regId, d);
  return json({ success:true, regId, message:'Your request has been submitted! We will contact you within 24 hours.' });
}
function getRegistrations(token) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const data=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registrations')?.getDataRange().getValues()||[];
  const headers=data[0]||[];
  return json({ success:true, registrations: data.slice(1).filter(r=>r[0]).map(r=>{
    const row={}; headers.forEach((h,j)=>row[h]=r[j]); return row;
  })});
}
function approveRegistration(token, regId) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const s=ss.getSheetByName('Registrations');
  const data=s.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(regId)){
      // Create customer account
      const custSheet=ss.getSheetByName('Customers');
      const custId='CUST'+Utilities.formatDate(new Date(),'Asia/Kolkata','yyMMddHHmmss');
      const pin=generatePin_();
      custSheet.appendRow([custId, data[i][1], data[i][2], data[i][3], data[i][4],
        data[i][7], data[i][5], 'Gold', data[i][6], 0,
        sha256_(pin), 'TRUE', nowIST_(), '', 'From: self-registration '+regId]);
      // Update registration status
      s.getRange(i+1,11).setValue('Approved'); s.getRange(i+1,12).setValue(nowIST_());
      // Send welcome email with PIN
      if(data[i][3]) sendRegistrationApprovalEmail_(data[i][1], data[i][2], data[i][3], custId, pin);
      return json({ success:true, customerId:custId, pin, emailSent:!!data[i][3] });
    }
  }
  return json({ success:false, error:'Registration not found' });
}
function rejectRegistration(token, regId, reason) {
  if (!validateToken_(token,'admin')) return json({ success:false, error:'Unauthorised' });
  const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registrations');
  const data=s.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===String(regId)){
      s.getRange(i+1,11).setValue('Rejected'); s.getRange(i+1,12).setValue(nowIST_());
      s.getRange(i+1,13).setValue(reason||'Outside delivery area');
      if(data[i][3]) sendRegistrationRejectionEmail_(data[i][1], data[i][3], reason);
      return json({ success:true });
    }
  }
  return json({ success:false, error:'Registration not found' });
}

// ==================== EMAIL BUILDER (Outlook + Webmail safe) ====================
// All emails use table-based layout, inline styles, no CSS3, bgcolor attributes
function emailWrap_(bodyHtml) {
  const biz   = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const city  = cfg_('BUSINESS_CITY') || 'Hubballi';
  const phone = cfg_('BUSINESS_PHONE') || '';
  const addr  = cfg_('BUSINESS_ADDRESS') || '';
  const upi   = cfg_('UPI_ID') || '';

  return `<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->
<style type="text/css">
  body{margin:0;padding:0;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}
  table,td{mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;}
  img{border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic;}
  a{color:#1a6b8a;text-decoration:none;}
  @media only screen and (max-width:600px){
    .email-container{width:100%!important;}
    .email-padding{padding:16px!important;}
  }
</style>
</head>
<body style="margin:0;padding:0;background-color:#f0f4f8;">
<!--[if mso]><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td bgcolor="#f0f4f8" align="center"><![endif]-->
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f0f4f8;">
<tr><td align="center" style="padding:24px 8px;">
  <table class="email-container" width="580" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff;max-width:580px;">

    <!-- HEADER -->
    <tr>
      <td bgcolor="#1a6b8a" align="center" style="padding:28px 24px;background-color:#1a6b8a;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr><td align="center">
            <p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:26px;font-weight:bold;color:#ffffff;mso-line-height-rule:exactly;line-height:32px;">&#x1F95B; ${biz}</p>
            <p style="margin:6px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#b3dae8;mso-line-height-rule:exactly;line-height:18px;">${city} &nbsp;&#124;&nbsp; Fresh Milk, Doorstep Daily</p>
          </td></tr>
        </table>
      </td>
    </tr>

    <!-- BODY -->
    <tr>
      <td class="email-padding" style="padding:28px 32px;font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333333;line-height:22px;">
        ${bodyHtml}
      </td>
    </tr>

    <!-- FOOTER -->
    <tr>
      <td bgcolor="#1a6b8a" style="padding:18px 24px;background-color:#1a6b8a;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr><td align="center" style="font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#b3dae8;line-height:20px;">
            <strong style="color:#ffffff;">${biz}</strong><br>
            ${phone ? `&#128222; ${phone}` : ''}${addr ? `&nbsp; &#124; &nbsp;${addr}` : ''}<br>
            ${upi ? `<span style="color:#ffffff;">UPI: <strong>${upi}</strong></span>` : ''}
          </td></tr>
        </table>
      </td>
    </tr>

  </table>
</td></tr>
</table>
<!--[if mso]></td></tr></table><![endif]-->
</body></html>`;
}

function emailDivider_() {
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0;">
    <tr><td bgcolor="#e8f4f8" style="height:1px;font-size:0;line-height:0;">&nbsp;</td></tr>
  </table>`;
}

function emailDetailTable_(rows) {
  // rows = array of [label, value, isHighlight]
  let inner = '';
  rows.forEach(([label, value, highlight]) => {
    const valStyle = highlight
      ? 'font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:bold;color:#1a6b8a;padding:9px 14px;border-bottom:1px solid #e8f4f8;'
      : 'font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333333;padding:9px 14px;border-bottom:1px solid #e8f4f8;';
    inner += `<tr>
      <td width="40%" style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#7f8c8d;padding:9px 14px;border-bottom:1px solid #e8f4f8;white-space:nowrap;">${label}</td>
      <td width="60%" style="${valStyle}">${value}</td>
    </tr>`;
  });
  return `<table width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;border:1px solid #bfd8e3;">
    <tr><td bgcolor="#1a6b8a" colspan="2" style="background-color:#1a6b8a;padding:10px 14px;font-family:Arial,Helvetica,sans-serif;font-size:13px;font-weight:bold;color:#ffffff;">DETAILS</td></tr>
    ${inner}
  </table>`;
}

function emailTotalBox_(label, amount, cur) {
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;">
    <tr>
      <td bgcolor="#1a6b8a" style="background-color:#1a6b8a;padding:14px 18px;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:16px;font-weight:bold;color:#ffffff;">${label}</td>
            <td align="right" style="font-family:Arial,Helvetica,sans-serif;font-size:22px;font-weight:bold;color:#ffffff;">${cur}${Number(amount).toLocaleString('en-IN')}</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>`;
}

function emailUpiBox_(upi, billId) {
  if (!upi) return '';
  return `<table width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;border:1px solid #a5d6a7;margin:16px 0;">
    <tr><td bgcolor="#e8f5e9" align="center" style="background-color:#e8f5e9;padding:16px 20px;">
      <p style="margin:0;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:#1b5e20;">Pay via UPI &nbsp;&#x2192;&nbsp; <span style="font-size:17px;">${upi}</span></p>
      <p style="margin:6px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#388e3c;">GPay &nbsp;&#124;&nbsp; PhonePe &nbsp;&#124;&nbsp; Paytm &nbsp;&#124;&nbsp; Any UPI App</p>
      ${billId ? `<p style="margin:6px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#555;">Remarks: ${billId}</p>` : ''}
    </td></tr>
  </table>`;
}

function emailAlertBox_(text, bgColor, txtColor) {
  return `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:14px 0;">
    <tr><td bgcolor="${bgColor}" style="background-color:${bgColor};padding:14px 18px;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:${txtColor};line-height:20px;">
      ${text}
    </td></tr>
  </table>`;
}

function dryRunNote_(originalEmail) {
  if (!isDryRun_()) return '';
  return emailAlertBox_(`&#x1F9EA; <strong>DRY RUN MODE</strong> &mdash; Original recipient: ${originalEmail}`, '#fff3cd', '#856404');
}

function sendEmail_(to, subject, bodyHtml) {
  if (!to) return false;
  const biz = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  try {
    MailApp.sendEmail({ to, subject, htmlBody: emailWrap_(bodyHtml), name: biz, replyTo: cfg_('SUPPORT_EMAIL') || cfg_('ADMIN_EMAIL') || '' });
    return true;
  } catch(e) { Logger.log('Email error: ' + e); return false; }
}

// ==================== CREDENTIAL EMAIL ====================
function sendCredentialEmail_(custId, name, phone, pin, email) {
  const to  = resolveEmail_(email);
  if (!to) return false;
  const url = cfg_('GITHUB_PAGES_URL');
  const biz = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const body = `
    ${dryRunNote_(email)}
    <p style="margin:0 0 16px;">Dear <strong>${name}</strong>,</p>
    <p style="margin:0 0 20px;">Welcome to <strong>${biz}</strong>! Your delivery account has been created. Use the details below to log in to your customer portal.</p>
    <table width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;border:2px solid #1a6b8a;margin:20px 0;">
      <tr><td bgcolor="#1a6b8a" style="background-color:#1a6b8a;padding:12px 16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:#ffffff;">&#x1F511; YOUR LOGIN CREDENTIALS</td></tr>
      <tr><td style="padding:20px 16px;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#7f8c8d;padding-bottom:10px;">&#128222; Mobile Number</td>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:bold;color:#333;padding-bottom:10px;">${phone}</td>
          </tr>
          <tr>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#7f8c8d;vertical-align:middle;">&#x1F510; 6-Digit PIN</td>
            <td style="font-family:Arial,Helvetica,sans-serif;font-size:32px;font-weight:bold;color:#1a6b8a;letter-spacing:10px;">${pin}</td>
          </tr>
        </table>
        ${url ? `<p style="margin:14px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;">&#x1F310; Portal: <a href="${url}login.html" style="color:#1a6b8a;">${url}login.html</a></p>` : ''}
      </td></tr>
    </table>
    ${emailAlertBox_('&#x26A0;&#xFE0F; Please keep your PIN confidential. Do not share it with anyone.', '#fff8e1', '#e65100')}
    <p style="margin:12px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">Use the portal to view your daily delivery status, track your milk calendar, and see monthly bills.</p>`;
  return sendEmail_(to, `Your ${biz} Portal PIN: ${pin}`, body);
}

// ==================== BILL EMAIL ====================
function sendBillEmail_(bill) {
  const to  = resolveEmail_(bill.email);
  if (!to) return false;
  const cur = cfg_('CURRENCY_SYMBOL') || '₹';
  const upi = cfg_('UPI_ID') || '';
  const biz = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const rows = [
    ['Bill Reference', bill.billId, false],
    ['Billing Period', `${monthName_(bill.month)} ${bill.year}`, false],
    ['Days Delivered', `${bill.totalDays} days`, false],
    ['Milk', `${parseFloat(bill.milkLtrs).toFixed(1)} Ltrs &nbsp;&#x2192;&nbsp; ${cur}${Number(bill.milkAmount).toLocaleString('en-IN')}`, false],
  ];
  if (bill.curdLtrs > 0) rows.push(['Curd', `${parseFloat(bill.curdLtrs).toFixed(1)} Ltrs &nbsp;&#x2192;&nbsp; ${cur}${Number(bill.curdAmount).toLocaleString('en-IN')}`, false]);
  if (bill.deliveryCharge > 0) rows.push(['Delivery Charge', `${cur}${bill.deliveryCharge}`, false]);
  let detailTable = emailDetailTable_(rows);
  // Override heading in detail table
  detailTable = detailTable.replace('>DETAILS<', `>MONTHLY BILL &mdash; ${monthName_(bill.month).toUpperCase()} ${bill.year}<`);
  const body = `
    ${dryRunNote_(bill.email)}
    <p style="margin:0 0 16px;">Dear <strong>${bill.name}</strong>,</p>
    <p style="margin:0 0 20px;">Please find your milk delivery bill for <strong>${monthName_(bill.month)} ${bill.year}</strong> below.</p>
    ${detailTable}
    ${emailTotalBox_('TOTAL AMOUNT DUE', bill.totalAmount, cur)}
    ${emailUpiBox_(upi, bill.billId)}
    <p style="margin:16px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">For any queries, reply to this email or call us directly. Thank you for choosing ${biz}!</p>`;
  return sendEmail_(to, `${biz} Bill — ${monthName_(bill.month)} ${bill.year} — ${cur}${bill.totalAmount}`, body);
}

// ==================== PAYMENT RECEIPT EMAIL ====================
function sendPaymentReceiptEmail_(p) {
  const to  = resolveEmail_(p.custEmail);
  if (!to) return false;
  const cur = cfg_('CURRENCY_SYMBOL') || '₹';
  const biz = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const rows = [
    ['Receipt No.',        p.payId,                       true],
    ['Payment Date',       nowIST_().split(' ')[0],        false],
    ['Amount Received',    `${cur}${Number(p.amount).toLocaleString('en-IN')}`, true],
    ['Payment Method',     p.method,                      false],
    ['Transaction / Ref',  p.transactionRef || 'N/A',     false],
    ['Billing Period',     `${monthName_(p.month)} ${p.year}`, false],
  ];
  if (p.billId) rows.push(['Bill Reference', p.billId, false]);
  let detailTable = emailDetailTable_(rows);
  detailTable = detailTable.replace('>DETAILS<', '>PAYMENT RECEIPT<');
  const body = `
    ${dryRunNote_(p.custEmail)}
    <p style="margin:0 0 16px;">Dear <strong>${p.custName}</strong>,</p>
    <p style="margin:0 0 20px;">&#x2705; Thank you! We have received your payment. Please find your receipt below for your records.</p>
    ${detailTable}
    ${emailAlertBox_(`&#x2705; <strong>Payment of ${cur}${Number(p.amount).toLocaleString('en-IN')} confirmed via ${p.method}.</strong><br>No further action needed.`, '#d4edda', '#155724')}
    <p style="margin:12px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">Please keep this email as your payment proof. Thank you for choosing ${biz}!</p>`;
  return sendEmail_(to, `${biz} — Payment Receipt ${p.payId} — ${cur}${Number(p.amount).toLocaleString('en-IN')}`, body);
}

// ==================== REGISTRATION EMAILS ====================
function sendRegistrationAckEmail_(name, email) {
  const to  = resolveEmail_(email);
  if (!to) return false;
  const biz  = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const phone = cfg_('BUSINESS_PHONE') || '';
  const body = `
    ${dryRunNote_(email)}
    <p style="margin:0 0 16px;">Dear <strong>${name}</strong>,</p>
    <p style="margin:0 0 16px;">&#x1F44B; Thank you for your interest in <strong>${biz}</strong>!</p>
    <p style="margin:0 0 20px;">We have received your delivery request. Our team will review your details and confirm if we can deliver to your area.</p>
    ${emailAlertBox_('&#x23F0; <strong>What happens next?</strong><br>Our team will contact you within <strong>24 hours</strong> to confirm your delivery or call you for more details.', '#d1ecf1', '#0c5460')}
    <p style="margin:16px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">If you need to reach us urgently, call us at <strong>${phone}</strong>.</p>
    <p style="margin:8px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">Thank you for your patience. We look forward to serving you fresh milk every day!</p>`;
  return sendEmail_(to, `${biz} — Delivery Request Received — We'll be in touch soon!`, body);
}

function sendRegistrationAdminAlert_(regId, d) {
  const adminEmail = cfg_('ADMIN_EMAIL');
  if (!adminEmail) return;
  const body = `
    <p style="margin:0 0 16px;"><strong>New customer registration received:</strong></p>
    ${emailDetailTable_([
      ['Reg ID', regId, true],
      ['Name', d.name, false],
      ['Phone', d.phone, false],
      ['Email', d.email || 'Not provided', false],
      ['Address', d.address || 'Not provided', false],
      ['Preferred Brand', d.preferredBrand || 'Amul', false],
      ['Preferred Qty', (d.preferredQty || 1) + ' Ltrs/day', false],
      ['Shift', d.shift || 'Morning', false],
      ['Message', d.message || '—', false],
    ])}
    ${emailAlertBox_('&#x26A0;&#xFE0F; Please log in to the Admin Panel &rarr; Registrations tab to Approve or Reject this request.', '#fff3cd', '#856404')}`;
  try { MailApp.sendEmail({ to: adminEmail, subject: `[NEW REGISTRATION] ${d.name} — ${d.phone}`, htmlBody: emailWrap_(body), name: 'Milk Parlour System' }); } catch(e) { Logger.log(e); }
}

function sendRegistrationApprovalEmail_(name, phone, email, custId, pin) {
  const to  = resolveEmail_(email);
  if (!to) return false;
  const biz = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const url = cfg_('GITHUB_PAGES_URL') || '';
  const body = `
    ${dryRunNote_(email)}
    <p style="margin:0 0 16px;">Dear <strong>${name}</strong>,</p>
    ${emailAlertBox_('&#x1F389; <strong>Great news! Your delivery request has been APPROVED!</strong>', '#d4edda', '#155724')}
    <p style="margin:16px 0;">We are delighted to confirm that <strong>${biz}</strong> will be delivering fresh milk to your doorstep. Your account has been created — here are your login details:</p>
    <table width="100%" cellpadding="0" cellspacing="0" border="1" style="border-collapse:collapse;border:2px solid #1a6b8a;margin:16px 0;">
      <tr><td bgcolor="#1a6b8a" style="background-color:#1a6b8a;padding:12px 16px;font-family:Arial,Helvetica,sans-serif;font-size:14px;font-weight:bold;color:#ffffff;">&#x1F511; YOUR PORTAL LOGIN</td></tr>
      <tr><td style="padding:20px 16px;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
          <tr><td style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#7f8c8d;padding-bottom:8px;">&#128222; Mobile Number</td><td style="font-family:Arial,Helvetica,sans-serif;font-size:15px;font-weight:bold;color:#333;padding-bottom:8px;">${phone}</td></tr>
          <tr><td style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#7f8c8d;">&#x1F510; 6-Digit PIN</td><td style="font-family:Arial,Helvetica,sans-serif;font-size:32px;font-weight:bold;color:#1a6b8a;letter-spacing:10px;">${pin}</td></tr>
        </table>
        ${url ? `<p style="margin:14px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;">&#x1F310; Login at: <a href="${url}login.html" style="color:#1a6b8a;">${url}login.html</a></p>` : ''}
      </td></tr>
    </table>
    ${emailAlertBox_('&#x26A0;&#xFE0F; Please keep your PIN confidential. Do not share it with anyone.', '#fff8e1', '#e65100')}
    <p style="margin:12px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">Your delivery will begin from the next working day. Welcome to the ${biz} family! &#x1F95B;</p>`;
  return sendEmail_(to, `${biz} — Your delivery is CONFIRMED! Login PIN inside`, body);
}

function sendRegistrationRejectionEmail_(name, email, reason) {
  const to  = resolveEmail_(email);
  if (!to) return false;
  const biz   = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const phone = cfg_('BUSINESS_PHONE') || '';
  const body = `
    ${dryRunNote_(email)}
    <p style="margin:0 0 16px;">Dear <strong>${name}</strong>,</p>
    <p style="margin:0 0 16px;">Thank you for your interest in <strong>${biz}</strong>.</p>
    <p style="margin:0 0 16px;">Unfortunately, we are unable to fulfil your delivery request at this time.</p>
    ${reason ? emailAlertBox_(`<strong>Reason:</strong> ${reason}`, '#f8d7da', '#721c24') : ''}
    <p style="margin:16px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">We are continuously expanding our delivery network. Please feel free to reach out again in the future — we would love to serve you!</p>
    ${phone ? `<p style="margin:8px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">For more information, call us at <strong>${phone}</strong>.</p>` : ''}
    <p style="margin:16px 0 0;font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#555;">We are sorry for the inconvenience and thank you for your understanding.</p>`;
  return sendEmail_(to, `${biz} — Update on your delivery request`, body);
}

// ==================== WHATSAPP ====================
function sendBillWhatsApp_(bill) {
  const phone = resolvePhone_(bill.phone);
  if (!phone||cfg_('WHATSAPP_PROVIDER')==='none') return false;
  const cur = cfg_('CURRENCY_SYMBOL')||'₹'; const upi = cfg_('UPI_ID')||'';
  const msg = `🥛 *${cfg_('BUSINESS_NAME')||'Manish Milk Parlour'}*\n\nDear ${bill.name},\nYour bill for *${monthName_(bill.month)} ${bill.year}*:\n\n📦 Days: ${bill.totalDays}\n🥛 Milk: ${parseFloat(bill.milkLtrs).toFixed(1)}L — ${cur}${bill.milkAmount}${bill.curdLtrs>0?`\n🍦 Curd: ${parseFloat(bill.curdLtrs).toFixed(1)}L — ${cur}${bill.curdAmount}`:''}${bill.deliveryCharge>0?`\n🚚 Delivery: ${cur}${bill.deliveryCharge}`:''}\n\n*TOTAL: ${cur}${bill.totalAmount}*${upi?`\n\nPay via UPI: *${upi}*\nGPay / PhonePe / Paytm`:''}\n\nBill ID: ${bill.billId}`;
  try { return sendViaTwilio_(phone, msg); } catch(e) { Logger.log(e); return false; }
}
function sendPaymentReceiptWhatsApp_(p) {
  const phone = resolvePhone_(p.custPhone);
  if (!phone||cfg_('WHATSAPP_PROVIDER')==='none') return false;
  const cur = cfg_('CURRENCY_SYMBOL')||'₹';
  const msg = `🥛 *${cfg_('BUSINESS_NAME')||'Manish Milk Parlour'}*\n\n✅ *Payment Receipt*\n\nDear ${p.custName},\nWe have received your payment.\n\n📋 Receipt: ${p.payId}\n💰 Amount: *${cur}${Number(p.amount).toLocaleString('en-IN')}*\n💳 Method: ${p.method}${p.transactionRef?`\n🔖 Ref: ${p.transactionRef}`:''}\n📅 Period: ${monthName_(p.month)} ${p.year}\n\nThank you! ✨`;
  try { return sendViaTwilio_(phone, msg); } catch(e) { Logger.log(e); return false; }
}
function sendViaTwilio_(toPhone, msg) {
  const sid=cfg_('WHATSAPP_ACCOUNT_SID'), tok=cfg_('WHATSAPP_AUTH_TOKEN'), from=cfg_('WHATSAPP_FROM')||'whatsapp:+14155238886';
  if (!sid||!tok) return false;
  const to=`whatsapp:+91${String(toPhone).replace(/\D/g,'').slice(-10)}`;
  const res=UrlFetchApp.fetch(`https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`,{
    method:'post', headers:{Authorization:'Basic '+Utilities.base64Encode(`${sid}:${tok}`)},
    payload:{From:from, To:to, Body:msg}
  });
  return res.getResponseCode()===201;
}

// ==================== SEED SAMPLE DATA (Run once for demo/dry-run) ====================
function seedFromExcelData() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const btn = ui.alert('Seed customer & delivery data from Milk Order-2023.xlsx?\n\n' +
    'This will ADD 21 customers and October 2023 delivery records.\n' +
    'Safe to run on an empty sheet — will skip if CustomerID already exists.',
    ui.ButtonSet.OK_CANCEL);
  if (btn !== ui.Button.OK) return;

  const custSheet  = ss.getSheetByName('Customers');
  const delivSheet = ss.getSheetByName('Deliveries');
  if (!custSheet || !delivSheet) { ui.alert('Run setupMilkSheets() first.'); return; }

  // ── 21 customers from Milk Order-2023.xlsx ──────────────────────
  const customers = [
    { id:'CUST001', name:'Manish G Khona',              address:'KDO Manison B#203',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST002', name:'Vijaykumar Bisnahalli',        address:'Manoj Paradise 4th Cross', shift:'Morning', brand:'Amul+Aditya',  quality:'Gold',               milkQty:1,   curdQty:0.2 },
    { id:'CUST003', name:'Rashmi Munavar',               address:'KDO Residency A#1',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST004', name:'Mayur R Lodaya',               address:'KDO Residency A#2',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:2,   curdQty:0   },
    { id:'CUST005', name:'Sandeep H Lodaya',             address:'KDO Residency A#3',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:2,   curdQty:0   },
    { id:'CUST006', name:'Shantilal R Lalka',            address:'KDO Residency A#4',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST007', name:'Bharat H Khona',               address:'KDO Residency A#5',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST008', name:'Padam D Nagda',                address:'KDO Residency A#6',        shift:'Morning', brand:'Amul',         quality:'Taza',               milkQty:1,   curdQty:0   },
    { id:'CUST009', name:'Bharati S Momaya',             address:'KDO Residency A#10',       shift:'Morning', brand:'Amul',         quality:'Gold+Taza',          milkQty:1.5, curdQty:0   },
    { id:'CUST010', name:'Uday Kurwa',                   address:'KDO Residency A#11',       shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST011', name:'Sachin Maisheri',              address:'KDO Residency A#12',       shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST012', name:'Jayesh G Kohna',               address:'KDO Residency A#14',       shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
    { id:'CUST013', name:'Jitin S Momaya',               address:'KDO Residency A#16',       shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1.5, curdQty:0   },
    { id:'CUST014', name:'Yellamma Gudi',                address:'Security House',            shift:'Morning', brand:'Nandini',      quality:'Pasteurised Toned',  milkQty:0.5, curdQty:0   },
    { id:'CUST015', name:'Laxmichand R Lodaya',          address:'KDO Residency B#1',        shift:'Morning', brand:'Kolhapur',     quality:'Regular',            milkQty:1,   curdQty:0   },
    { id:'CUST016', name:'Vimala C Dharamshi',           address:'KDO Residency B#12',       shift:'Morning', brand:'Kolhapur',     quality:'Regular',            milkQty:1,   curdQty:0   },
    { id:'CUST017', name:'Laxmiben H Sheth',             address:'KDO Residency B#3',        shift:'Morning', brand:'Kolhapur',     quality:'Regular',            milkQty:1,   curdQty:0   },
    { id:'CUST018', name:'Mukesh Valji Poladia',         address:'KDO Residency B#4',        shift:'Morning', brand:'Nandini+Taza', quality:'Pasteurised Toned',  milkQty:1.5, curdQty:0   },
    { id:'CUST019', name:'Chandrakant J Shah',           address:'KDO Residency B#6',        shift:'Morning', brand:'Kolhapur',     quality:'Nandini+Gold',       milkQty:1.5, curdQty:0   },
    { id:'CUST020', name:'Kaushik K Dharmshi',           address:'KDO Residency B#7',        shift:'Morning', brand:'Amul',         quality:'Taza',               milkQty:1.5, curdQty:0   },
    { id:'CUST021', name:'Jignesh J Lodaya',             address:'KDO Residency B#8',        shift:'Morning', brand:'Amul',         quality:'Gold',               milkQty:1,   curdQty:0   },
  ];

  // Oct 2023 delivery days per customer (True = delivered, from Excel)
  // Index = day-1 (0=day1 … 30=day31). Only days 11-15 had data in Oct sheet.
  const octDeliveries = {
    'CUST001': [11,12,13,14],
    'CUST002': [11,12,13,14,15],
    'CUST003': [14,15],
    'CUST004': [11,12,13,14],
    'CUST005': [12],
    'CUST006': [11,12,13,14,15],
    'CUST007': [11,12,13,14,15],
    'CUST008': [11,13,14,15],
    'CUST009': [11,12,13,14,15],
    'CUST010': [11,12,13,14,15],
    'CUST011': [13,14],
    'CUST012': [13,14,15],
    'CUST013': [],
    'CUST014': [11,12,13,14,15],
    'CUST015': [11,12,13,14,15],
    'CUST016': [12,14],
    'CUST017': [11,12,13,14,15],
    'CUST018': [11,12,13,14,15],
    'CUST019': [11,12,13,14,15],
    'CUST020': [11,12,13,14,15],
    'CUST021': [11,12,13,14,15],
  };

  // Check existing customer IDs to avoid duplicates
  const existingCustData = custSheet.getDataRange().getValues();
  const existingIds = new Set(existingCustData.slice(1).map(r => String(r[0])));

  let custAdded = 0, delivAdded = 0;
  const now = nowIST_();

  customers.forEach(c => {
    if (!existingIds.has(c.id)) {
      const pin = generatePin_();
      custSheet.appendRow([
        c.id, c.name, '', '', c.address, c.shift,
        c.brand, c.quality, c.milkQty, c.curdQty,
        sha256_(pin), 'TRUE', now, '', 'Seeded from Milk Order-2023.xlsx'
      ]);
      custAdded++;
    }
  });

  // Seed October 2023 delivery records
  const existDeliv = delivSheet.getDataRange().getValues();
  const existDelivKeys = new Set(existDeliv.slice(1).map(r => `${r[1]}_${r[3]}`));

  customers.forEach(c => {
    const days = octDeliveries[c.id] || [];
    days.forEach(day => {
      const dateStr = `2023-10-${String(day).padStart(2,'0')}`;
      const key = `${c.id}_${dateStr}`;
      if (existDelivKeys.has(key)) return;
      const delivId = `DEL-OCT23-${c.id}-D${day}`;
      delivSheet.appendRow([
        delivId, c.id, c.name, dateStr, '07:00:00',
        day, 10, 2023, 'SEED', 'Seeded',
        c.brand, c.quality, c.milkQty, c.curdQty,
        'Delivered', 'Seeded from Excel'
      ]);
      delivAdded++;
    });
  });

  ui.alert(
    `✅ Seed complete!\n\n` +
    `👥 Customers added: ${custAdded} (${customers.length - custAdded} already existed)\n` +
    `📦 Deliveries added: ${delivAdded} (October 2023)\n\n` +
    `Next: Admin Panel → Bills → Select October 2023 → Generate Bills → Send Bills\n\n` +
    `Note: PINs are auto-generated. Use "Send PIN" in Admin to email them, or use Admin Panel → Customers to view.`
  );
}

function seedFullDemoMonth() {
  // Seeds a COMPLETE month (all 31 days delivered) for all customers
  // Useful for a demo invoice showing full month billing
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const btn = ui.alert(
    'Seed FULL DEMO month (October 2023)?\n\n' +
    'This seeds all 21 customers with deliveries on ALL 31 days of October 2023.\n' +
    'Useful to generate a realistic full-month demo invoice.\n' +
    'Run seedFromExcelData() first to ensure customers exist.',
    ui.ButtonSet.OK_CANCEL);
  if (btn !== ui.Button.OK) return;

  const custSheet  = ss.getSheetByName('Customers');
  const delivSheet = ss.getSheetByName('Deliveries');
  const custData   = custSheet.getDataRange().getValues();

  const existDeliv = delivSheet.getDataRange().getValues();
  const existDelivKeys = new Set(existDeliv.slice(1).map(r => `${r[1]}_${r[3]}`));

  let count = 0;
  for (let i = 1; i < custData.length; i++) {
    if (!custData[i][0] || custData[i][11] !== 'TRUE') continue;
    const custId   = String(custData[i][0]);
    const custName = custData[i][1];
    const brand    = custData[i][6], quality = custData[i][7];
    const milkQty  = parseFloat(custData[i][8]) || 0;
    const curdQty  = parseFloat(custData[i][9]) || 0;

    for (let day = 1; day <= 31; day++) {
      const dateStr = `2023-10-${String(day).padStart(2,'0')}`;
      const key = `${custId}_${dateStr}`;
      if (existDelivKeys.has(key)) continue;
      delivSheet.appendRow([
        `DEL-OCT23-FULL-${custId}-D${day}`, custId, custName, dateStr, '07:00:00',
        day, 10, 2023, 'DEMO', 'Demo seed',
        brand, quality, milkQty, curdQty, 'Delivered', 'Full demo month'
      ]);
      count++;
    }
  }

  ui.alert(`✅ Full demo month seeded!\n\n📦 ${count} delivery records added for October 2023.\n\nNow go to Admin Panel → Bills → October 2023 → Generate Bills.`);
}

// ==================== EOD ORDER SYSTEM ====================

function tomorrowDateIST_() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}

function isCutoffPassed_() {
  const cutoff = cfg_('ORDER_CUTOFF_TIME') || '19:00';
  const [ch, cm] = cutoff.split(':').map(Number);
  const now = new Date();
  const nowIST = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Kolkata' }));
  return nowIST.getHours() > ch || (nowIST.getHours() === ch && nowIST.getMinutes() >= cm);
}

function placeTomorrowOrder(token, d) {
  const sess = validateToken_(token, 'customer');
  if (!sess) return json({ success: false, error: 'Unauthorised' });
  if (isCutoffPassed_()) {
    const cutoff = cfg_('ORDER_CUTOFF_TIME') || '19:00';
    return json({ success: false, error: `Order cutoff of ${cutoff} has passed. Your default order will be delivered tomorrow.` });
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TomorrowOrders');
  const custSheet = ss.getSheetByName('Customers');
  const custData = custSheet?.getDataRange().getValues() || [];
  let custName = '', custPhone = '';
  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][0]) === String(sess.id)) { custName = custData[i][1]; custPhone = custData[i][2]; break; }
  }
  const forDate = tomorrowDateIST_();
  const now = nowIST_();
  const existing = sheet.getDataRange().getValues();
  for (let i = 1; i < existing.length; i++) {
    if (String(existing[i][1]) === String(sess.id) && String(existing[i][4]) === forDate) {
      sheet.getRange(i+1, 6).setValue(d.brand);
      sheet.getRange(i+1, 7).setValue(d.quality || '');
      sheet.getRange(i+1, 8).setValue(d.milkQty || 0);
      sheet.getRange(i+1, 9).setValue(d.curdQty || 0);
      sheet.getRange(i+1, 11).setValue(now);
      return json({ success: true, updated: true, forDate, brand: d.brand, milkQty: d.milkQty, curdQty: d.curdQty });
    }
  }
  const orderId = 'ORD' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMddHHmmss');
  sheet.appendRow([orderId, sess.id, custName, custPhone, forDate, d.brand, d.quality || '', d.milkQty || 0, d.curdQty || 0, now, '']);
  return json({ success: true, updated: false, forDate, brand: d.brand, milkQty: d.milkQty, curdQty: d.curdQty });
}

function getTomorrowOrder(token) {
  const sess = validateToken_(token, 'customer');
  if (!sess) return json({ success: false, error: 'Unauthorised' });
  const cutoff = cfg_('ORDER_CUTOFF_TIME') || '19:00';
  const cutoffPassed = isCutoffPassed_();
  const forDate = tomorrowDateIST_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('TomorrowOrders');
  const custSheet = ss.getSheetByName('Customers');
  // Get customer defaults
  const custData = custSheet?.getDataRange().getValues() || [];
  let defaults = {};
  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][0]) === String(sess.id)) {
      defaults = { brand: custData[i][6], quality: custData[i][7], milkQty: custData[i][8], curdQty: custData[i][9] };
      break;
    }
  }
  // Check for a custom order for tomorrow
  const data = sheet?.getDataRange().getValues() || [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sess.id) && String(data[i][4]) === forDate) {
      return json({ success: true, forDate, cutoffPassed, cutoffTime: cutoff,
        order: { brand: data[i][5], quality: data[i][6], milkQty: data[i][7], curdQty: data[i][8], isModified: true, placedAt: data[i][9] },
        defaults });
    }
  }
  return json({ success: true, forDate, cutoffPassed, cutoffTime: cutoff, order: null, defaults });
}

function getWholesalePrices(token) {
  if (!validateToken_(token, 'admin')) return json({ success: false, error: 'Unauthorised' });
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WholesalePrices')?.getDataRange().getValues() || [];
  return json({ success: true, prices: data.slice(1).filter(r => r[0]).map(r => ({ brand: r[0], quality: r[1], wholesalePrice: r[2] })) });
}

function updateWholesalePrice(token, d) {
  if (!validateToken_(token, 'admin')) return json({ success: false, error: 'Unauthorised' });
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WholesalePrices');
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === d.brand && data[i][1] === d.quality) {
      s.getRange(i+1, 3).setValue(d.wholesalePrice);
      s.getRange(i+1, 4).setValue(nowIST_());
      return json({ success: true });
    }
  }
  s.appendRow([d.brand, d.quality, d.wholesalePrice, nowIST_()]);
  return json({ success: true });
}

function getEODSummary(token) {
  if (!validateToken_(token, 'admin') && !validateToken_(token, 'delivery')) return json({ success: false, error: 'Unauthorised' });
  const forDate = tomorrowDateIST_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('TomorrowOrders');
  const custSheet  = ss.getSheetByName('Customers');
  const wpSheet    = ss.getSheetByName('WholesalePrices');
  const custData   = custSheet?.getDataRange().getValues() || [];
  const orderData  = orderSheet?.getDataRange().getValues() || [];
  const wpData     = wpSheet?.getDataRange().getValues() || [];
  // Build wholesale price map
  const wpMap = {};
  for (let i = 1; i < wpData.length; i++) {
    if (wpData[i][0]) wpMap[`${wpData[i][0]}|${wpData[i][1]}`] = parseFloat(wpData[i][2]) || 0;
  }
  // Build map of custom orders for tomorrow
  const customOrders = {};
  for (let i = 1; i < orderData.length; i++) {
    if (String(orderData[i][4]) === forDate) customOrders[String(orderData[i][1])] = orderData[i];
  }
  // Aggregate: for each active customer, use custom order if placed, else default
  const brandTotals = {}; // key: "brand|quality" → { brand, quality, totalMilkQty, totalCurdQty, customers: [] }
  let grandMilkQty = 0, grandCurdQty = 0, grandCost = 0;
  for (let i = 1; i < custData.length; i++) {
    if (!custData[i][0] || custData[i][11] !== 'TRUE') continue;
    const custId = String(custData[i][0]);
    let brand, quality, milkQty, curdQty;
    if (customOrders[custId]) {
      brand = customOrders[custId][5]; quality = customOrders[custId][6];
      milkQty = parseFloat(customOrders[custId][7]) || 0; curdQty = parseFloat(customOrders[custId][8]) || 0;
    } else {
      brand = custData[i][6]; quality = custData[i][7];
      milkQty = parseFloat(custData[i][8]) || 0; curdQty = parseFloat(custData[i][9]) || 0;
    }
    const key = `${brand}|${quality}`;
    if (!brandTotals[key]) brandTotals[key] = { brand, quality, milkQty: 0, curdQty: 0, customerCount: 0 };
    brandTotals[key].milkQty += milkQty;
    brandTotals[key].curdQty += curdQty;
    brandTotals[key].customerCount++;
    grandMilkQty += milkQty;
    grandCurdQty += curdQty;
    const wp = wpMap[key] || wpMap[`${brand}|`] || 0;
    grandCost += (milkQty + curdQty) * wp;
  }
  const summary = Object.values(brandTotals).sort((a, b) => (b.milkQty + b.curdQty) - (a.milkQty + a.curdQty)).map(r => ({
    ...r,
    milkQty: Math.round(r.milkQty * 100) / 100,
    curdQty: Math.round(r.curdQty * 100) / 100,
    wholesalePrice: wpMap[`${r.brand}|${r.quality}`] || 0,
    totalCost: Math.round((r.milkQty + r.curdQty) * (wpMap[`${r.brand}|${r.quality}`] || 0)),
  }));
  return json({ success: true, forDate, summary,
    grandMilkQty: Math.round(grandMilkQty * 100) / 100,
    grandCurdQty: Math.round(grandCurdQty * 100) / 100,
    grandTotalQty: Math.round((grandMilkQty + grandCurdQty) * 100) / 100,
    grandCost: Math.round(grandCost),
    customOrderCount: Object.keys(customOrders).length,
    cutoffPassed: isCutoffPassed_(),
  });
}

function sendEODReport(token) {
  if (!validateToken_(token, 'admin')) return json({ success: false, error: 'Unauthorised' });
  const wholesalerPhone = cfg_('WHOLESALER_PHONE');
  if (!wholesalerPhone) return json({ success: false, error: 'WHOLESALER_PHONE not set in Config sheet.' });
  const res = JSON.parse(getEODSummary(token).getContent());
  if (!res.success) return json(res);
  const { summary, forDate, grandTotalQty, grandCost, customOrderCount } = res;
  const biz = cfg_('BUSINESS_NAME') || 'Manish Milk Parlour';
  const wsName = cfg_('WHOLESALER_NAME') || 'Supplier';
  const cur = cfg_('CURRENCY_SYMBOL') || '₹';
  const d = new Date(forDate);
  const dateStr = d.toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long' });
  let lines = summary.map(r => {
    const qty = r.milkQty + (r.curdQty > 0 ? ` milk + ${r.curdQty} curd` : '');
    const cost = r.wholesalePrice > 0 ? ` @ ${cur}${r.wholesalePrice}/L = ${cur}${r.totalCost}` : '';
    return `• ${r.brand} ${r.quality}: *${r.milkQty}L*${r.curdQty > 0 ? ` + ${r.curdQty}L curd` : ''}${cost}`;
  }).join('\n');
  const msg =
    `🥛 *${biz}*\n` +
    `📦 *Order for ${dateStr}*\n\n` +
    `Hello ${wsName},\nKindly arrange the following:\n\n` +
    `${lines}\n\n` +
    `━━━━━━━━━━━━━━━\n` +
    `🛒 *Total Milk: ${res.grandMilkQty}L*${res.grandCurdQty > 0 ? `  |  Curd: ${res.grandCurdQty}L` : ''}\n` +
    `📦 *Grand Total: ${grandTotalQty}L*\n` +
    (grandCost > 0 ? `💰 *Est. Amount: ${cur}${grandCost.toLocaleString('en-IN')}*\n` : '') +
    `━━━━━━━━━━━━━━━\n` +
    `${customOrderCount > 0 ? `_(${customOrderCount} customer(s) modified today's order)_\n` : ''}` +
    `_Thank you_`;
  const sent = sendViaTwilio_(wholesalerPhone, msg);
  // Log in admin sheet (reuse AdminSheet row 3+)
  const adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin');
  try { adminSheet.appendRow([`eod_report_${forDate}`, nowIST_(), `Sent to ${wholesalerPhone}: ${grandTotalQty}L, ${cur}${grandCost}`]); } catch(_) {}
  return json({ success: true, sent, message: msg, forDate, grandTotalQty, grandCost });
}

// ==================== GOOGLE SHEET MENU ====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🥛 Milk Parlour')
    .addItem('🔧 Setup / Refresh Sheets', 'setupMilkSheets')
    .addSeparator()
    .addItem('📋 Generate Bills — Current Month', 'menuGenerateBills_')
    .addItem('📤 Send Bills — Current Month', 'menuSendBills_')
    .addSeparator()
    .addItem('📜 Pending Registrations', 'menuShowRegistrations_')
    .addSeparator()
    .addItem('🛒 View EOD Order Summary', 'menuEODSummary_')
    .addItem('📲 Send EOD Order to Wholesaler (WhatsApp)', 'menuSendEOD_')
    .addSeparator()
    .addItem('🔗 Show Web App URL', 'menuShowUrl_')
    .addItem('♻️ How to Recover Deleted Rows', 'menuRecoveryHelp_')
    .addSeparator()
    .addItem('🧪 Seed Customers from Excel (Oct 2023 partial)', 'seedFromExcelData')
    .addItem('📅 Seed Full Demo Month (Oct 2023 all 31 days)', 'seedFullDemoMonth')
    .addToUi();
}
function menuGenerateBills_() {
  const {month:m,year:y}=monthYearIST_();
  const t=issueToken_('admin','admin');
  const r=JSON.parse(generateBills(t,m,y).getContent());
  SpreadsheetApp.getUi().alert(`✅ Generated ${r.count} bills for ${monthName_(m)} ${y}\nTotal: ₹${r.generated.reduce((s,b)=>s+b.totalAmount,0).toLocaleString('en-IN')}`);
}
function menuSendBills_() {
  const {month:m,year:y}=monthYearIST_();
  const t=issueToken_('admin','admin');
  const r=JSON.parse(sendBills(t,m,y,null).getContent());
  const dry=isDryRun_()?'\n\n[DRY RUN — emails sent to admin only]':'';
  SpreadsheetApp.getUi().alert(`✅ Sent ${r.count} bills.${dry}`);
}
function menuShowRegistrations_() {
  const t=issueToken_('admin','admin');
  const r=JSON.parse(getRegistrations(t).getContent());
  const pending=r.registrations.filter(x=>x.Status==='Pending');
  SpreadsheetApp.getUi().alert(pending.length?`${pending.length} pending registrations:\n\n${pending.map(x=>`• ${x.Name} — ${x.Phone} (${x.RegisteredAt})`).join('\n')}\n\nUse Admin Portal → Registrations tab to approve/reject.`:'No pending registrations.');
}
function menuShowUrl_() {
  SpreadsheetApp.getUi().alert('Web App URL:\n\n' + ScriptApp.getService().getUrl() + '\n\nCopy this into:\n1. milk-app.js → GAS_URL variable\n2. Config sheet → GITHUB_PAGES_URL');
}
function menuEODSummary_() {
  const t = issueToken_('admin','admin');
  const r = JSON.parse(getEODSummary(t).getContent());
  if (!r.success) { SpreadsheetApp.getUi().alert('Error: ' + r.error); return; }
  const cur = cfg_('CURRENCY_SYMBOL') || '₹';
  const lines = r.summary.map(s =>
    `• ${s.brand} ${s.quality}: ${s.milkQty}L${s.curdQty > 0 ? ' + '+s.curdQty+'L curd' : ''} ${s.wholesalePrice > 0 ? '@ '+cur+s.wholesalePrice+' = '+cur+s.totalCost : ''}`
  ).join('\n');
  SpreadsheetApp.getUi().alert(
    `📦 EOD Order Summary for ${r.forDate}\n\n${lines}\n\n` +
    `TOTAL: ${r.grandTotalQty}L${r.grandCost > 0 ? '  |  Est. Cost: '+cur+r.grandCost.toLocaleString('en-IN') : ''}\n` +
    `(${r.customOrderCount} customers modified today)`
  );
}
function menuSendEOD_() {
  const phone = cfg_('WHOLESALER_PHONE');
  if (!phone) { SpreadsheetApp.getUi().alert('⚠️ Set WHOLESALER_PHONE in Config sheet first.'); return; }
  if (!SpreadsheetApp.getUi().alert('Send EOD order to wholesaler ' + phone + '?',
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) === SpreadsheetApp.getUi().Button.OK) return;
  const t = issueToken_('admin','admin');
  const r = JSON.parse(sendEODReport(t).getContent());
  SpreadsheetApp.getUi().alert(r.success
    ? `✅ EOD order sent to ${phone}!\n\nTotal: ${r.grandTotalQty}L`
    : `❌ Failed: ${r.error}`);
}
function menuRecoveryHelp_() {
  SpreadsheetApp.getUi().alert(
    '♻️ How to Recover Accidentally Deleted Rows\n\n' +
    '1. Go to: File → Version History → See version history\n' +
    '2. Browse timestamps to find when the data existed\n' +
    '3. Click "Restore this version"\n\n' +
    'Alternatively:\n' +
    '• Ctrl+Z (Undo) works immediately after deletion\n' +
    '• All sheets have a warning prompt before editing protected ranges\n\n' +
    'TIP: Google Sheets keeps 30 days of version history automatically.'
  );
}
