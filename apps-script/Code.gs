/* ===== Lefkada 2027 Booking System — Code.gs ===== */

const SHEET_CABINS = 'Cabins';
const SHEET_BOOKINGS = 'Bookings';
const SHEET_WAITLIST = 'Waitlist';
const SHEET_CONFIG = 'Config';
const SHEET_PAYMENTS = 'Payments';
const SHEET_LOG = 'Log';

// Boat 1 = Bali 5.2 (6 cabins, bookable now). Boat 2 = Lagoon 46 (4 cabins, waitlist
// until Config!boat2_open = TRUE). Any non-'bali' boat is treated as the second boat.
const CABIN_TYPES = [
  { boat: 'bali', type: 'double', label: 'Bali 5.2 - Double Cabin', priceTotal: 9000, guests: 2, ids: ['BALI-C1', 'BALI-C2', 'BALI-C3', 'BALI-C4', 'BALI-C5', 'BALI-C6'] },
  { boat: 'lagoon', type: 'double', label: 'Lagoon 46 - Double Cabin', priceTotal: 9000, guests: 2, ids: ['LAGOON-C1', 'LAGOON-C2', 'LAGOON-C3', 'LAGOON-C4'] }
];

const CONFIG_DEFAULTS = {
  admin_email: 'info@sailing2wellness.com',
  partner_email: 'yogawithlorraine13@gmail.com',
  boat2_open: 'FALSE',
  boat2_threshold: '6',
  reservation_expiry_hours: '48',
  card_surcharge_pct: '3.5',
  inst1_flat_per_guest: '1000',
  inst2_pct_of_remainder: '25',
  inst3_pct_of_remainder: '25',
  inst4_pct_of_remainder: '50',
  inst2_due_date: '2026-11-01',
  inst3_due_date: '2027-03-01',
  inst4_due_date: '2027-06-26',
  trip_departure_date: '2027-09-18',
  last_poll_timestamp: '0',
  waitlist_threshold_notified: 'FALSE',
  terms_url: 'https://sailing2wellness.com/terms/',
  bank_name: 'Sailing2Wellness',
  bank_routing: '084009519',
  bank_account_number: '252003195652603',
  bank_swift: 'TRWIUS35XXX',
  bank_address: 'Wise US Inc, 108 W 13th St, Wilmington, DE, 19801, United States',
  bank_currency: 'USD'
};

/* ===== One-run setup functions ===== */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const cabinsSheet = getOrCreateSheet(ss, SHEET_CABINS,
    ['cabin_id', 'boat', 'type', 'label', 'price_pp', 'price_total', 'status', 'booking_id']);
  if (cabinsSheet.getLastRow() < 2) {
    const rows = [];
    CABIN_TYPES.forEach(ct => {
      const pricePp = ct.priceTotal / ct.guests;
      ct.ids.forEach(id => {
        const status = ct.boat === 'bali' ? 'available' : 'hidden';
        rows.push([id, ct.boat, ct.type, ct.label, pricePp, ct.priceTotal, status, '']);
      });
    });
    cabinsSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  getOrCreateSheet(ss, SHEET_BOOKINGS,
    ['booking_id', 'timestamp', 'cabin_id', 'name', 'email', 'phone', 'guests', 'guest2_name',
     'status', 'inst1_paid', 'inst2_paid', 'inst3_paid', 'inst4_paid', 'notes']);

  getOrCreateSheet(ss, SHEET_WAITLIST,
    ['timestamp', 'name', 'email', 'phone', 'guests', 'boat2_interest']);

  const configSheet = getOrCreateSheet(ss, SHEET_CONFIG, ['key', 'value']);
  const existingKeys = configSheet.getLastRow() > 1
    ? configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues().flat()
    : [];
  const newRows = [];
  Object.keys(CONFIG_DEFAULTS).forEach(key => {
    if (existingKeys.indexOf(key) === -1) newRows.push([key, CONFIG_DEFAULTS[key]]);
  });
  if (newRows.length) configSheet.getRange(configSheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);

  getOrCreateSheet(ss, SHEET_PAYMENTS,
    ['timestamp', 'stripe_session_id', 'booking_id', 'amount', 'installment']);

  getOrCreateSheet(ss, SHEET_LOG, ['timestamp', 'event', 'detail']);

  logEvent('setup', 'setupSheet completed');
}

function createStripeLinks() {
  const props = PropertiesService.getScriptProperties();
  const setupKey = props.getProperty('STRIPE_SETUP_KEY');
  if (!setupKey) throw new Error('Set STRIPE_SETUP_KEY in Script Properties (Project Settings) before running this.');

  const cardPct = getConfig('card_surcharge_pct');
  if (cardPct === '' || cardPct === null || cardPct === undefined) {
    throw new Error('Set Config!card_surcharge_pct (a number, e.g. 3.5) before running createStripeLinks(). Run setupSheet() first if the key is missing.');
  }

  CABIN_TYPES.forEach(ct => {
    const installments = computeInstallments(ct.priceTotal, ct.guests);
    installments.forEach((amount, idx) => {
      const instNum = idx + 1;
      const finalAmount = round2(amount * (1 + Number(cardPct) / 100));
      const url = createOnePaymentLink(setupKey, ct, instNum, finalAmount);
      setConfig(`link_${ct.boat}_${ct.type}_inst${instNum}_card`, url);
    });
  });

  logEvent('setup', 'createStripeLinks completed');
}

function createOnePaymentLink(key, ct, instNum, amount) {
  const priceRes = stripeApiCall(key, 'prices', {
    unit_amount: Math.round(amount * 100),
    currency: 'eur',
    'product_data[name]': `${ct.label} - Installment ${instNum} of 4 (Card)`
  });
  const linkRes = stripeApiCall(key, 'payment_links', {
    'line_items[0][price]': priceRes.id,
    'line_items[0][quantity]': 1,
    'metadata[boat]': ct.boat,
    'metadata[cabin_type]': ct.type,
    'metadata[installment]': String(instNum),
    'metadata[payment_method]': 'card'
  });
  return linkRes.url;
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('pollStripePayments').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('expireReservations').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendInstallmentReminders').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('weeklyOutstandingReport').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).create();
  ScriptApp.newTrigger('checkWaitlistThreshold').timeBased().everyDays(1).atHour(8).create();

  logEvent('setup', 'installTriggers completed - 5 triggers created');
}

/* ===== Web app entry points ===== */

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'availability') return jsonOut(getAvailability());
    if (action === 'healthcheck') return jsonOut(getHealthcheck());
    return jsonOut({ ok: false, reason: 'unknown_action' });
  } catch (err) {
    logEvent('error', 'doGet: ' + err.message);
    emailAdmin('Booking system error (doGet)', err.message);
    return jsonOut({ ok: false, reason: 'server_error' });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'reserve') return jsonOut(handleReserve(body));
    if (body.action === 'waitlist') return jsonOut(handleWaitlist(body));
    return jsonOut({ ok: false, reason: 'unknown_action' });
  } catch (err) {
    logEvent('error', 'doPost: ' + err.message);
    emailAdmin('Booking system error (doPost)', err.message);
    return jsonOut({ ok: false, reason: 'server_error' });
  }
}

function getAvailability() {
  const rows = readSheet(SHEET_CABINS);
  const boat2Open = getConfig('boat2_open') === 'TRUE';
  const cabins = rows
    .filter(r => r.boat === 'bali' || boat2Open || r.status !== 'hidden')
    .map(r => ({
      cabin_id: r.cabin_id, boat: r.boat, type: r.type, label: r.label,
      price_pp: r.price_pp, price_total: r.price_total, status: r.status,
      installments: computeInstallments(Number(r.price_total), r.type === 'single' ? 1 : 2)
    }));
  return {
    cabins,
    terms_url: getConfig('terms_url'),
    card_surcharge_pct: Number(getConfig('card_surcharge_pct') || 0),
    bank: {
      name: getConfig('bank_name'),
      routing: getConfig('bank_routing'),
      account_number: getConfig('bank_account_number'),
      swift: getConfig('bank_swift'),
      address: getConfig('bank_address'),
      currency: getConfig('bank_currency')
    }
  };
}

function getHealthcheck() {
  const logRows = readSheet(SHEET_LOG);
  const lastLog = logRows.length ? logRows[logRows.length - 1].timestamp : null;
  return {
    ok: true,
    triggers: ScriptApp.getProjectTriggers().length,
    last_poll_timestamp: getConfig('last_poll_timestamp'),
    last_log_timestamp: lastLog
  };
}

/* ===== Reservation flow ===== */

function handleReserve(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const err = validateReservation(body);
    if (err) return { ok: false, reason: err };

    const existing = findBooking(body.booking_id);
    if (existing) return { ok: true, booking_id: existing.booking_id, idempotent: true };

    const cabin = findCabin(body.cabin_id);
    if (!cabin) return { ok: false, reason: 'cabin_not_found' };
    if (cabin.boat !== 'bali' && getConfig('boat2_open') !== 'TRUE') return { ok: false, reason: 'boat_not_open' };
    if (cabin.status !== 'available') return { ok: false, reason: 'cabin_taken' };

    const method = body.payment_method === 'bank' ? 'bank' : 'card';
    const status = method === 'bank' ? 'pending_bank' : 'pending';

    appendRow(SHEET_BOOKINGS, [
      body.booking_id, new Date().toISOString(), body.cabin_id, body.name, body.email,
      body.phone || '', body.guests, body.guest2_name || '', status,
      'FALSE', 'FALSE', 'FALSE', 'FALSE', method
    ]);
    updateCabinStatus(body.cabin_id, 'reserved', body.booking_id);

    const guests = cabin.type === 'single' ? 1 : 2;
    const installments = computeInstallments(Number(cabin.price_total), guests);

    if (method === 'bank') {
      sendGuestEmail(body.email, 'Your Lefkada 2027 reservation - bank transfer details',
        bankReservationEmailBody(body, cabin, installments));
      emailAdmin('New reservation (BANK TRANSFER): ' + body.name,
        `Cabin ${body.cabin_id} (${cabin.label}), ${body.guests} guest(s). Booking ID ${body.booking_id}. Watch your Wise USD account for the deposit, then mark inst1_paid=TRUE and status=confirmed in the Bookings tab.`);
    } else {
      const links = getLinksForCabin(cabin, 1, body.booking_id, body.email);
      sendGuestEmail(body.email, 'Your Lefkada 2027 reservation - next step',
        reservationEmailBody(body, cabin, installments, links));
      emailAdmin('New reservation: ' + body.name,
        `Cabin ${body.cabin_id} (${cabin.label}), ${body.guests} guest(s). Booking ID ${body.booking_id}.`);
    }
    logEvent('reserve', `${body.booking_id} -> ${body.cabin_id} (${method})`);

    return { ok: true, booking_id: body.booking_id, method, installments };
  } finally {
    lock.releaseLock();
  }
}

function validateReservation(body) {
  if (!body.booking_id) return 'missing_booking_id';
  if (!body.name) return 'missing_name';
  if (!body.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(body.email)) return 'invalid_email';
  const guests = Number(body.guests);
  if (!guests || guests < 1 || guests > 2) return 'invalid_guests';
  if (!body.cabin_id) return 'missing_cabin_id';
  return null;
}

function handleWaitlist(body) {
  if (!body.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(body.email)) return { ok: false, reason: 'invalid_email' };
  appendRow(SHEET_WAITLIST, [new Date().toISOString(), body.name || '', body.email, body.phone || '', body.guests || '', body.boat2_interest || 'TRUE']);
  sendGuestEmail(body.email, "You're on the Lefkada 2027 waitlist",
    `Hi ${body.name || ''},<br><br>You're on the waitlist for Boat 2. We'll email you the moment it's confirmed.<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`);
  emailPartner('New Lagoon 46 (Boat 2) waitlist signup', `${body.name} (${body.email}) joined the Lagoon 46 waitlist.`);
  logEvent('waitlist', body.email);
  return { ok: true };
}

/* ===== Triggers ===== */

function expireReservations() {
  // Only auto-expires CARD reservations ('pending'). Bank-transfer holds ('pending_bank')
  // never auto-expire, since wires take days - Vincent confirms or cancels those by hand.
  const expiryHours = Number(getConfig('reservation_expiry_hours'));
  const rows = readSheetWithRowIndex(SHEET_BOOKINGS);
  const now = new Date();
  rows.forEach(({ row, index }) => {
    if (row.status !== 'pending') return;
    const ageHours = (now - new Date(row.timestamp)) / 36e5;
    if (ageHours < expiryHours) return;
    setCellValue(SHEET_BOOKINGS, index, 'status', 'expired');
    updateCabinStatus(row.cabin_id, 'available', '');
    sendGuestEmail(row.email, 'Your Lefkada 2027 reservation has expired',
      `Hi ${row.name},<br><br>Your reservation on ${row.cabin_id} expired after ${expiryHours} hours without a deposit. The cabin is available again - feel free to book another if you'd still like to join.<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`);
    emailAdmin('Reservation expired', `${row.name} - ${row.cabin_id} (booking ${row.booking_id})`);
    logEvent('expire', row.booking_id);
  });
}

function pollStripePayments() {
  const key = PropertiesService.getScriptProperties().getProperty('STRIPE_POLLING_KEY');
  if (!key) { logEvent('error', 'STRIPE_POLLING_KEY not set'); return; }

  const lastPoll = Number(getConfig('last_poll_timestamp') || '0');
  const nowTs = Math.floor(Date.now() / 1000);
  const resp = stripeApiGet(key, 'checkout/sessions', { limit: 100, 'created[gte]': lastPoll });

  (resp.data || []).forEach(session => {
    if (session.payment_status !== 'paid') return;
    const bookingId = session.client_reference_id;
    const instNum = session.metadata && session.metadata.installment;
    if (!bookingId || !instNum) return;
    if (paymentAlreadyRecorded(session.id)) return;
    recordPayment(session.id, bookingId, session.amount_total / 100, instNum);
  });

  setConfig('last_poll_timestamp', String(nowTs));
  logEvent('poll', `checked ${(resp.data || []).length} sessions since ${lastPoll}`);
}

function recordPayment(sessionId, bookingId, amount, instNum) {
  appendRow(SHEET_PAYMENTS, [new Date().toISOString(), sessionId, bookingId, amount, instNum]);

  const booking = findBooking(bookingId);
  if (!booking) { logEvent('error', `payment for unknown booking ${bookingId}`); return; }

  setBookingField(bookingId, `inst${instNum}_paid`, 'TRUE');

  if (String(instNum) === '1') {
    setBookingField(bookingId, 'status', 'confirmed');
    updateCabinStatus(booking.cabin_id, 'sold', bookingId);
    sendGuestEmail(booking.email, "You're confirmed! Lefkada 2027",
      `Hi ${booking.name},<br><br>Your deposit is in - you're confirmed on ${booking.cabin_id} for The Wellness Odyssey, Sep 18-24 2027. We'll email you ahead of each remaining installment.<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`);
  } else {
    sendGuestEmail(booking.email, `Payment received - installment ${instNum}`,
      `Hi ${booking.name},<br><br>Received your installment ${instNum} payment of EUR ${amount} ${usdApprox(amount)}. Thank you.<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`);
  }
  emailAdmin('Payment received', `${booking.name}, ${booking.cabin_id}, installment ${instNum}, EUR ${amount}`);
  logEvent('payment', `${bookingId} inst${instNum} EUR ${amount}`);
}

function sendInstallmentReminders() {
  const rows = readSheet(SHEET_BOOKINGS);
  const today = new Date();
  const dueDates = { 2: getConfig('inst2_due_date'), 3: getConfig('inst3_due_date'), 4: getConfig('inst4_due_date') };
  const cardPct = Number(getConfig('card_surcharge_pct') || 0);

  rows.filter(r => r.status === 'confirmed').forEach(booking => {
    const method = booking.notes === 'bank' ? 'bank' : 'card';
    [2, 3, 4].forEach(n => {
      if (booking[`inst${n}_paid`] === 'TRUE') return;
      const due = new Date(dueDates[n]);
      const daysUntilDue = Math.floor((due - today) / 86400000);
      const isReminderDay = daysUntilDue === 14 || daysUntilDue === 0 || (daysUntilDue < 0 && Math.abs(daysUntilDue) % 5 === 0);
      if (!isReminderDay) return;
      if (reminderAlreadySentToday(booking.booking_id, n)) return;

      const cabin = findCabin(booking.cabin_id);
      const baseAmt = computeInstallments(Number(cabin.price_total), cabin.type === 'single' ? 1 : 2)[n - 1];

      let payBlock;
      if (method === 'bank') {
        payBlock = `Installment ${n} for ${cabin.label} (EUR ${baseAmt} ${usdApprox(baseAmt)}) is due ${dueDates[n]}.<br><br>` +
          bankDetailsHtml() +
          `Reference: ${booking.booking_id}`;
      } else {
        const cardAmt = round2(baseAmt * (1 + cardPct / 100));
        const links = getLinksForCabin(cabin, n, booking.booking_id, booking.email);
        payBlock = `Installment ${n} for ${cabin.label} (EUR ${cardAmt} ${usdApprox(cardAmt)}, incl. ${cardPct}% card fee) is due ${dueDates[n]}.<br><br>` +
          `Pay by card: ${links.card}<br>Reference: ${booking.booking_id}`;
      }

      sendGuestEmail(booking.email, `Installment ${n} due - Lefkada 2027`,
        `Hi ${booking.name},<br><br>${payBlock}<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`);
      logEvent('reminder', `${booking.booking_id} inst${n}`);
    });
  });
}

function weeklyOutstandingReport() {
  const rows = readSheet(SHEET_BOOKINGS).filter(r => r.status === 'confirmed');
  const dueDates = { 2: getConfig('inst2_due_date'), 3: getConfig('inst3_due_date'), 4: getConfig('inst4_due_date') };
  const today = new Date();
  const lines = [];

  rows.forEach(booking => {
    const method = booking.notes === 'bank' ? 'bank' : 'card';
    [2, 3, 4].forEach(n => {
      if (booking[`inst${n}_paid`] === 'TRUE') return;
      const due = new Date(dueDates[n]);
      const daysOverdue = Math.floor((today - due) / 86400000);
      const cabin = findCabin(booking.cabin_id);
      const amount = computeInstallments(Number(cabin.price_total), cabin.type === 'single' ? 1 : 2)[n - 1];
      lines.push(`${booking.name} | ${booking.cabin_id} | ${method} | Installment ${n} | EUR ${amount} | ${daysOverdue > 0 ? daysOverdue + ' days overdue' : 'not yet due'}`);
    });
  });

  const body = lines.length ? lines.join('<br>') : 'All payments up to date ✅';
  emailAdmin('Weekly outstanding payments - Lefkada 2027', body);
  logEvent('report', `${lines.length} outstanding line(s)`);
}

function checkWaitlistThreshold() {
  if (getConfig('boat2_open') === 'TRUE') return;
  if (getConfig('waitlist_threshold_notified') === 'TRUE') return;
  const count = readSheet(SHEET_WAITLIST).length;
  if (count < Number(getConfig('boat2_threshold'))) return;
  const msg = `${count} people are on the Lagoon 46 waitlist - enough interest to add the second boat.`;
  emailPartner('Your Lagoon 46 waitlist is ready', msg + ' We will get the second boat opened for booking.');
  emailAdmin('Boat 2 threshold reached', msg + ' Confirm a Lagoon 46 and set Config!boat2_open = TRUE to open booking.');
  setConfig('waitlist_threshold_notified', 'TRUE');
  logEvent('waitlist_threshold', String(count));
}

/* ===== Installments & links ===== */

function computeInstallments(priceTotal, guests) {
  const inst1 = Number(getConfig('inst1_flat_per_guest')) * guests;
  const remainder = priceTotal - inst1;
  const p2 = Number(getConfig('inst2_pct_of_remainder')) / 100;
  const p3 = Number(getConfig('inst3_pct_of_remainder')) / 100;
  const inst2 = round2(remainder * p2);
  const inst3 = round2(remainder * p3);
  const inst4 = round2(remainder - inst2 - inst3);
  return [inst1, inst2, inst3, inst4];
}

function getLinksForCabin(cabin, instNum, bookingId, email) {
  const rawCard = getConfig(`link_${cabin.boat}_${cabin.type}_inst${instNum}_card`);
  return { card: withBookingRef(rawCard, bookingId, email) };
}

function withBookingRef(url, bookingId, email) {
  const sep = url.includes('?') ? '&' : '?';
  return url + sep + 'client_reference_id=' + encodeURIComponent(bookingId) + '&prefilled_email=' + encodeURIComponent(email);
}

/* ===== Stripe REST helpers ===== */

function stripeApiCall(key, path, formParams) {
  const stringParams = {};
  Object.keys(formParams).forEach(k => stringParams[k] = String(formParams[k]));
  const options = {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(key + ':') },
    payload: stringParams,
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch('https://api.stripe.com/v1/' + path, options);
  const json = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() >= 300) throw new Error('Stripe error on ' + path + ': ' + resp.getContentText());
  return json;
}

function stripeApiGet(key, path, params) {
  const query = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
  const options = {
    method: 'get',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(key + ':') },
    muteHttpExceptions: true
  };
  const resp = UrlFetchApp.fetch('https://api.stripe.com/v1/' + path + '?' + query, options);
  const json = JSON.parse(resp.getContentText());
  if (resp.getResponseCode() >= 300) throw new Error('Stripe error on ' + path + ': ' + resp.getContentText());
  return json;
}

function paymentAlreadyRecorded(sessionId) {
  return readSheet(SHEET_PAYMENTS).some(r => r.stripe_session_id === sessionId);
}

/* ===== Sheet helpers ===== */

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function readSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).filter(r => r.join('') !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function readSheetWithRowIndex(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const out = [];
  values.forEach((row, i) => {
    if (i === 0 || row.join('') === '') return;
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    out.push({ row: obj, index: i + 1 });
  });
  return out;
}

function appendRow(name, values) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).appendRow(values);
}

function setCellValue(sheetName, rowIndex, columnName, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = headers.indexOf(columnName) + 1;
  sheet.getRange(rowIndex, col).setValue(value);
}

function findCabin(cabinId) {
  return readSheet(SHEET_CABINS).find(r => r.cabin_id === cabinId) || null;
}

function findBooking(bookingId) {
  return readSheet(SHEET_BOOKINGS).find(r => r.booking_id === bookingId) || null;
}

function updateCabinStatus(cabinId, status, bookingId) {
  const rows = readSheetWithRowIndex(SHEET_CABINS);
  const match = rows.find(r => r.row.cabin_id === cabinId);
  if (!match) return;
  setCellValue(SHEET_CABINS, match.index, 'status', status);
  setCellValue(SHEET_CABINS, match.index, 'booking_id', bookingId);
}

function setBookingField(bookingId, field, value) {
  const rows = readSheetWithRowIndex(SHEET_BOOKINGS);
  const match = rows.find(r => r.row.booking_id === bookingId);
  if (!match) return;
  setCellValue(SHEET_BOOKINGS, match.index, field, value);
}

function reminderAlreadySentToday(bookingId, instNum) {
  const today = new Date().toDateString();
  return readSheet(SHEET_LOG).some(r =>
    r.event === 'reminder' && r.detail === `${bookingId} inst${instNum}` && new Date(r.timestamp).toDateString() === today);
}

function getConfig(key) {
  const row = readSheet(SHEET_CONFIG).find(r => r.key === key);
  return row ? row.value : '';
}

function setConfig(key, value) {
  const rows = readSheetWithRowIndex(SHEET_CONFIG);
  const match = rows.find(r => r.row.key === key);
  if (match) {
    setCellValue(SHEET_CONFIG, match.index, 'value', value);
  } else {
    appendRow(SHEET_CONFIG, [key, value]);
  }
}

function logEvent(event, detail) {
  appendRow(SHEET_LOG, [new Date().toISOString(), event, detail]);
}

/* ===== Email ===== */

const EMAIL_SIGNATURE_HTML = `<table cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse; font-family: Arial, Helvetica, sans-serif; max-width: 400px;">
  <tr>
    <td style="padding-right: 14px; vertical-align: middle; border-right: 2px solid #dddddd;">
      <img src="https://sailing2wellness.com/wp-content/uploads/2023/11/Logo-2023-font-282x300.png" alt="Sailing2Wellness" width="90" style="display: block; max-width: 90px; height: auto;">
    </td>
    <td style="padding-left: 14px; vertical-align: top;">
      <div style="font-weight: bold; font-size: 15px; color: #1a3c5a; line-height: 1.3;">The Sailing2Wellness Team</div>
      <div style="font-size: 12px; color: #444444; margin-top: 6px; line-height: 1.6;">
        <b style="color: #1a3c5a;">E</b> <a href="mailto:info@sailing2wellness.com" style="color: #1a73e8; text-decoration: none;">info@sailing2wellness.com</a><br>
        <b style="color: #1a3c5a;">W</b> <a href="https://sailing2wellness.com" style="color: #1a73e8; text-decoration: none;">sailing2wellness.com</a>
      </div>
      <div style="margin-top: 8px;">
        <a href="https://admin.trustindex.io/widget/logClick?pub_widget_id=f1a279f46efc53111606adab86c" style="text-decoration: none; display: inline-block;">
          <table cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse;">
            <tr>
              <td style="vertical-align: middle; padding-right: 6px;">
                <span style="display: inline-block; width: 16px; height: 16px; background-color: #00b67a; border-radius: 50%; text-align: center; line-height: 16px; font-size: 10px; color: white; font-weight: bold;">&#10003;</span>
              </td>
              <td style="vertical-align: middle;">
                <span style="font-size: 12px; font-weight: bold; color: #333333;">Trustindex</span>
                <span style="font-size: 11px; color: #888888; margin-left: 3px;">56 reviews</span><br>
                <span style="font-size: 16px; color: #fbbc04; letter-spacing: 1px; line-height: 1;">&#9733;&#9733;&#9733;&#9733;&#9733;</span>
              </td>
            </tr>
          </table>
        </a>
      </div>
    </td>
  </tr>
</table>`;

function bankDetailsHtml() {
  return `<b>Bank transfer details (${getConfig('bank_currency')}, via Wise):</b><br>` +
    `Account name: ${getConfig('bank_name')}<br>` +
    `Routing number (US ACH/wire): ${getConfig('bank_routing')}<br>` +
    `Account number: ${getConfig('bank_account_number')}<br>` +
    `SWIFT/BIC (from outside the US): ${getConfig('bank_swift')}<br>` +
    `Bank address: ${getConfig('bank_address')}<br><br>` +
    `Sending from a US bank? Use the routing + account number for a cheap domestic transfer. Please send the USD equivalent of the amount due, and add your booking reference in the note.<br><br>`;
}

function sendGuestEmail(to, subject, htmlBody) {
  MailApp.sendEmail({ to, subject, htmlBody });
}

function emailAdmin(subject, body) {
  MailApp.sendEmail({ to: getConfig('admin_email') || CONFIG_DEFAULTS.admin_email, subject, htmlBody: body });
}

function emailPartner(subject, body) {
  const to = getConfig('partner_email');
  if (!to) return;
  MailApp.sendEmail({ to, subject, htmlBody: body + '<br><br>Best,<br>' + EMAIL_SIGNATURE_HTML });
}

function reservationEmailBody(body, cabin, installments, links) {
  const cardPct = Number(getConfig('card_surcharge_pct') || 0);
  const c = installments.map(a => round2(a * (1 + cardPct / 100)));
  return `Hi ${body.name},<br><br>` +
    `Your reservation on ${cabin.label} (${cabin.cabin_id || body.cabin_id}) is held for ${getConfig('reservation_expiry_hours')} hours.<br><br>` +
    `Your installments (card, incl. ${cardPct}% card processing fee):<br>` +
    `Installment 1 (due now): EUR ${c[0]} ${usdApprox(c[0])}<br>` +
    `Installment 2: EUR ${c[1]} ${usdApprox(c[1])}<br>` +
    `Installment 3: EUR ${c[2]} ${usdApprox(c[2])}<br>` +
    `Installment 4 (final, 12 weeks before departure): EUR ${c[3]} ${usdApprox(c[3])}<br><br>` +
    `Pay installment 1 by card: ${links.card}<br><br>` +
    `Prefer to pay by bank transfer at the flat price (no card fee)? Just reply and we'll send you the details.<br><br>` +
    `Booking reference: ${body.booking_id}<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`;
}

function bankReservationEmailBody(body, cabin, installments) {
  return `Hi ${body.name},<br><br>` +
    `Your reservation on ${cabin.label} (${cabin.cabin_id || body.cabin_id}) is held. To confirm it, please send your first installment by bank transfer - flat price, no card fee.<br><br>` +
    `<b>Installment 1 (due now): EUR ${installments[0]} ${usdApprox(installments[0])}</b><br>` +
    `Installment 2: EUR ${installments[1]} ${usdApprox(installments[1])}<br>` +
    `Installment 3: EUR ${installments[2]} ${usdApprox(installments[2])}<br>` +
    `Installment 4 (final, 12 weeks before departure): EUR ${installments[3]} ${usdApprox(installments[3])}<br><br>` +
    bankDetailsHtml() +
    `<b>Booking reference: ${body.booking_id}</b> (please include this in your transfer note)<br><br>` +
    `Once your transfer lands we'll confirm your cabin by email.<br><br>Best,<br>${EMAIL_SIGNATURE_HTML}`;
}

/* ===== JSON response ===== */

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

const EUR_TO_USD = 1.14; // approximate, for guest reference only

function usdApprox(n) {
  return '(approx. $' + Math.round(Number(n) * EUR_TO_USD).toLocaleString('en-US') + ')';
}
