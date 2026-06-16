/**
 * Aquatic Paradise ISUP Rentals — Auto-Dispatch
 * ------------------------------------------------------------------
 * Runs every time the Google booking form is submitted. It:
 *   1. Reads the new booking.
 *   2. Emails the owner a clean dispatch ticket.
 *   3. (Optional) Sends the owner a WhatsApp via CallMeBot.
 *   4. Appends the booking to a "Dispatch Log" tab with a status column.
 *   5. Builds one-tap WhatsApp links to reply to the customer and to
 *      hand the job to a field crew member (Shamar / Dravin).
 *
 * No secrets live in this file. All config is read from Script
 * Properties (Project Settings -> Script properties), so this file is
 * safe to commit to a public repository.
 *
 * SETUP (one time):
 *   1. Open the Google Form's responses spreadsheet.
 *   2. Extensions -> Apps Script. Paste this file in (replace Code.gs).
 *   3. Run setupAPR() once and grant permissions when prompted.
 *      That sets default config and installs the submit trigger.
 *   4. Project Settings -> Script properties: fill in OWNER_EMAIL,
 *      OWNER_WHATSAPP, and (optional) CALLMEBOT_APIKEY.
 *   5. Run testNotify() to send yourself a sample ticket.
 * ------------------------------------------------------------------
 */

/** Default config written on first setup. Edit live values in Script Properties. */
var APR_DEFAULTS = {
  OWNER_EMAIL:      'delroystapleton908@gmail.com',  // where dispatch tickets are emailed
  OWNER_WHATSAPP:   '17844963447',                   // owner number, digits only, country code first
  BUSINESS_NAME:    'Aquatic Paradise ISUP Rentals',
  CREW_SHAMAR:      '',   // Shamar's WhatsApp (digits only) — fill to enable hand-off link
  CREW_DRAVIN:      '',   // Dravin's WhatsApp (digits only)
  CALLMEBOT_APIKEY: ''    // optional: enables auto WhatsApp to the owner (see guide)
};

function cfg_(key) {
  var p = PropertiesService.getScriptProperties();
  var v = p.getProperty(key);
  return (v === null || v === undefined) ? (APR_DEFAULTS[key] || '') : v;
}

/** One-time setup: seed properties + install the onFormSubmit trigger. */
function setupAPR() {
  var props = PropertiesService.getScriptProperties();
  Object.keys(APR_DEFAULTS).forEach(function (k) {
    if (props.getProperty(k) === null) props.setProperty(k, APR_DEFAULTS[k]);
  });

  // Avoid duplicate triggers.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ensureLogSheet_();
  Logger.log('APR auto-dispatch installed. Fill OWNER_EMAIL / OWNER_WHATSAPP in Script properties.');
}

/** Main handler — fires on every form submission. */
function onFormSubmit(e) {
  var data = parseSubmission_(e);
  var ticket = buildTicket_(data);

  // 1) Email the owner (always-on, zero external dependency).
  var owner = cfg_('OWNER_EMAIL');
  if (owner) {
    MailApp.sendEmail({
      to: owner,
      subject: '🌊 New booking — ' + (data.name || 'Guest') + ' · ' + (data.date || 'date TBC'),
      htmlBody: ticket.html
    });
  }

  // 2) Optional WhatsApp to the owner via CallMeBot.
  var apiKey = cfg_('CALLMEBOT_APIKEY');
  var ownerWa = cfg_('OWNER_WHATSAPP');
  if (apiKey && ownerWa) {
    try {
      var url = 'https://api.callmebot.com/whatsapp.php'
        + '?phone=' + encodeURIComponent(ownerWa)
        + '&text=' + encodeURIComponent(ticket.text)
        + '&apikey=' + encodeURIComponent(apiKey);
      UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch (err) {
      Logger.log('WhatsApp send failed: ' + err);
    }
  }

  // 3) Append to the Dispatch Log with a NEW status.
  logBooking_(data, ticket.replyLink);
}

/** Normalise a submission into a tidy object, regardless of exact field names. */
function parseSubmission_(e) {
  var nv = (e && e.namedValues) ? e.namedValues : {};
  var flat = {};
  Object.keys(nv).forEach(function (k) {
    flat[k] = Array.isArray(nv[k]) ? nv[k].join(', ') : nv[k];
  });

  function pick(rx) {
    var hit = Object.keys(flat).find(function (k) { return rx.test(k); });
    return hit ? flat[hit] : '';
  }

  return {
    raw:      flat,
    name:     pick(/name/i),
    phone:    pick(/whats\s?app|phone|number|tel/i),
    email:    pick(/e-?mail/i),
    gear:     pick(/gear|package|equipment|item/i),
    date:     pick(/date|day/i),
    time:     pick(/time|start|delivery/i),
    location: pick(/location|where|villa|hotel|address|beach/i),
    payment:  pick(/pay/i),
    notes:    pick(/note|request|special|comment/i),
    when:     (e && e.values && e.values[0]) ? e.values[0] : new Date().toString()
  };
}

/** Build the owner-facing dispatch ticket (HTML email + plain text + links). */
function buildTicket_(d) {
  var digits = (d.phone || '').replace(/[^0-9]/g, '');
  var waCustomer = digits
    ? 'https://wa.me/' + digits + '?text=' + encodeURIComponent(
        'Hi ' + (d.name || '').split(' ')[0] + ' — Aquatic Paradise here! We received your booking'
        + (d.gear ? ' for ' + d.gear : '') + (d.date ? ' on ' + d.date : '')
        + '. We’ll confirm your delivery shortly. 🏝')
      : '';

  var rows = [
    ['Guest', d.name], ['WhatsApp', d.phone], ['Email', d.email],
    ['Gear / Package', d.gear], ['Date', d.date], ['Time', d.time],
    ['Location', d.location], ['Payment', d.payment], ['Notes', d.notes]
  ].filter(function (r) { return r[1]; });

  var html = '<div style="font-family:Arial,sans-serif;max-width:560px">'
    + '<h2 style="color:#1F8FA0;margin:0 0 4px">🌊 New Booking — ' + cfg_('BUSINESS_NAME') + '</h2>'
    + '<p style="color:#667;margin:0 0 14px">Received ' + d.when + '</p>'
    + '<table style="border-collapse:collapse;width:100%">'
    + rows.map(function (r) {
        return '<tr>'
          + '<td style="padding:8px 10px;background:#f4eddd;font-weight:bold;border:1px solid #e5ddc9;width:150px">' + r[0] + '</td>'
          + '<td style="padding:8px 10px;border:1px solid #e5ddc9">' + String(r[1]).replace(/</g, '&lt;') + '</td>'
          + '</tr>';
      }).join('')
    + '</table>'
    + (waCustomer ? '<p style="margin:16px 0 0"><a href="' + waCustomer + '" '
        + 'style="background:#25D366;color:#fff;padding:11px 18px;border-radius:8px;text-decoration:none;font-weight:bold">'
        + '💬 Reply to guest on WhatsApp</a></p>' : '')
    + '<p style="color:#999;font-size:12px;margin-top:18px">Auto-dispatch · Aquatic Paradise · Bequia, SVG</p>'
    + '</div>';

  var text = '🌊 NEW BOOKING — ' + cfg_('BUSINESS_NAME') + '\n'
    + rows.map(function (r) { return r[0] + ': ' + r[1]; }).join('\n')
    + (waCustomer ? '\nReply: ' + waCustomer : '');

  return { html: html, text: text, replyLink: waCustomer };
}

/** Ensure a "Dispatch Log" sheet exists with headers. */
function ensureLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Dispatch Log');
  if (!sh) {
    sh = ss.insertSheet('Dispatch Log');
    sh.appendRow(['Received', 'Status', 'Guest', 'WhatsApp', 'Gear', 'Date', 'Time', 'Location', 'Payment', 'Notes', 'Reply link']);
    sh.getRange('1:1').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function logBooking_(d, replyLink) {
  var sh = ensureLogSheet_();
  sh.appendRow([
    new Date(), 'NEW', d.name, d.phone, d.gear, d.date, d.time, d.location, d.payment, d.notes, replyLink
  ]);
}

/** Send yourself a sample ticket to confirm everything is wired. */
function testNotify() {
  onFormSubmit({
    namedValues: {
      'Full Name': ['Test Guest'],
      'WhatsApp Number': ['+1 784 555 0123'],
      'Gear / Package': ['Paddle Board Full Day — $150 XCD'],
      'Appointment Date': ['2026-06-20'],
      'Delivery Time': ['09:00'],
      'Choose Your Location': ['Princess Margaret Beach'],
      'Payment Preference': ['Cash on Delivery'],
      'Notes / Special Requests': ['This is a test booking.']
    },
    values: [new Date().toString()]
  });
}
