// ============================================================
// GOOGLE APPS SCRIPT — Paste this in script.google.com
// Handles: Google Sheet + Email confirmation
// AI Moviemaking Workshop version
// ============================================================

// -- CONFIG ---------------------------------------------------
// SHEET_NAME is read from Sheet1!B1 in the spreadsheet.
// To change the active batch: just update that ONE cell — no script edit needed.
const CONFIG = {
  SHEET_ID:        '1Nct9qgBSCj59H81QaNCfHebAucSPZoMOSXwsUa6atBI',
  SENDER_NAME:     'Alex — AI Moviemaking Workshop',
  WORKSHOP_DATE:   '5th–6th September 2026 (Sat–Sun), 11:00 AM IST both days',
  WHATSAPP_GROUP:  'https://chat.whatsapp.com/IbaS5oahigG5WPeDhXX1GV?mode=gi_t',
  // Pick any random string, put the SAME string here and in the webhook URL's
  // ?secret=... query param when you set up the webhook in Razorpay Dashboard.
  // Apps Script web apps can't read custom request headers, so this is how we
  // confirm a POST actually came from us (via the URL) rather than checking
  // Razorpay's X-Razorpay-Signature header the normal way.
  RAZORPAY_WEBHOOK_SECRET: 'de4651dd322e0d01a3647fd56572b3d9',
  // Meta Conversions API — sends the Purchase event server-side the instant
  // payment is confirmed, so it doesn't depend on the browser pixel firing
  // (ad blockers, iOS privacy settings, closed tabs, etc).
  META_PIXEL_ID:   '1049945680950084',
  META_CAPI_TOKEN: 'EAAXZBZC4bTt40BSb7sKZAZC9oAfx6bmE5v5KsxoPretoZBRKaDlfIvzh6wkVfKBpoFRtRK8eANaBwZA0aZCpHxZAC9vjNwr7tWfV1fIgIgIiZBxWW73Bh9KDN1QvX9RguIJhdR83OBvYxYEVrRLkY5WxAZAZAUopw0MTtCCJGI9yXDTdasxkusnpN5d1RtT2qUmYQZDZD',
};

function getActiveBatch() {
  const name = SpreadsheetApp.openById(CONFIG.SHEET_ID)
    .getSheetByName('Sheet1').getRange('B1').getValue();
  if (!name) throw new Error('Active batch not set! Put the tab name in Sheet1 cell B1.');
  return name.toString().trim();
}
// ------------------------------------------------------------

function doPost(e) {
  try {
    // Razorpay webhook call — identified by the ?secret=... query param we put
    // on the webhook URL (see CONFIG.RAZORPAY_WEBHOOK_SECRET above).
    if (e.parameter && e.parameter.secret) {
      return handleRazorpayWebhook(e);
    }

    // Legacy JSON-POST path (not used by the live site — it posts via GET —
    // kept for backward compatibility).
    const data = JSON.parse(e.postData.contents);
    const { name, email, phone, paymentId, amount } = data;
    if (!name && !email && !phone) {
      return jsonOut({ success: false, error: 'no user data' });
    }
    saveLead(name, email, phone, paymentId, amount);
    if (paymentId && paymentId !== 'PAYMENT_INITIATED' && paymentId !== 'LEAD') {
      sendEmail(name, email, paymentId);
    }
    return jsonOut({ success: true });
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

// ------------------------------------------------------------
// RAZORPAY WEBHOOK — fires server-to-server the instant a payment is
// captured, regardless of whether the customer's browser ever makes it
// back to thankyou.html. This is what actually closes the "paid on
// Razorpay but sheet still says Initiated" gap.
// ------------------------------------------------------------
function handleRazorpayWebhook(e) {
  if (e.parameter.secret !== CONFIG.RAZORPAY_WEBHOOK_SECRET) {
    return jsonOut({ success: false, error: 'bad secret' });
  }
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.event !== 'payment_link.paid') {
      // Ack anything else so Razorpay doesn't keep retrying it.
      return jsonOut({ success: true, ignored: body.event });
    }

    const linkEntity = body.payload.payment_link.entity;
    const payEntity  = body.payload.payment.entity;
    const customer   = linkEntity.customer || {};

    const name       = customer.name    || payEntity.email || 'Customer';
    const email      = customer.email   || payEntity.email || '';
    const contact    = customer.contact || payEntity.contact || '';
    const amountPaid = linkEntity.amount_paid || payEntity.amount || 0;
    const isRecording = amountPaid >= 19900;
    const paymentId   = 'RAZORPAY_' + payEntity.id;

    const lock = LockService.getScriptLock();
    let result;
    try {
      lock.waitLock(15000);
      result = updateLeadStatus(contact || email, paymentId, amountPaid, isRecording);
      if (result === 'not_found') {
        saveLead(name, email, contact, paymentId, amountPaid, isRecording);
        result = 'updated';
      }
    } finally {
      lock.releaseLock();
    }
    // Only email/CAPI when this call actually flipped the row — avoids a
    // duplicate email if the browser's own thankyou.html callback already did it.
    if (result === 'updated' && email) {
      sendEmail(name, email, paymentId, isRecording);
      sendCAPIPurchase(payEntity.id, email, contact, isRecording);
    }
    return jsonOut({ success: true });
  } catch (err) {
    Logger.log('webhook error: ' + err.message);
    return jsonOut({ success: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// META CONVERSIONS API — server-side Purchase event
// ------------------------------------------------------------
function sha256Hex(input) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// rawPaymentId must be the bare Razorpay payment id (e.g. "pay_XXXX"), with
// no "RAZORPAY_" prefix — it's used as the event_id so Meta can de-duplicate
// this against the browser-pixel Purchase event fired on thankyou.html,
// which sends the same bare id as its eventID.
function sendCAPIPurchase(rawPaymentId, email, phone, isRecording) {
  try {
    const userData = {};
    if (email) userData.em = [sha256Hex(email.trim().toLowerCase())];
    if (phone) {
      let digits = String(phone).replace(/\D/g, '');
      if (digits.length === 10) digits = '91' + digits; // assume India if no country code
      userData.ph = [sha256Hex(digits)];
    }

    const payload = {
      data: [{
        event_name: 'Purchase',
        event_time: Math.floor(Date.now() / 1000),
        event_id: rawPaymentId,
        action_source: 'website',
        event_source_url: 'https://buildmyaimovie.in/thankyou.html',
        user_data: userData,
        custom_data: {
          value: isRecording ? 199 : 99,
          currency: 'INR',
          content_name: 'AI Moviemaking Workshop',
          content_type: 'product'
        }
      }]
    };

    const url = 'https://graph.facebook.com/v19.0/' + CONFIG.META_PIXEL_ID
      + '/events?access_token=' + CONFIG.META_CAPI_TOKEN;

    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log('CAPI response: ' + resp.getResponseCode() + ' ' + resp.getContentText());
  } catch (err) {
    Logger.log('CAPI error: ' + err.message);
  }
}

function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    if (p && p.name && p.email) {
      const isPaid = p.paymentId && p.paymentId !== 'LEAD' && p.paymentId !== 'PAYMENT_INITIATED';

      if (isPaid) {
        const rawAmount   = parseInt(p.amount) || 0;
        const isRecording = p.plan === 'recording' || rawAmount >= 19900;
        const paidAmount  = rawAmount > 0 ? rawAmount : (isRecording ? 19900 : 9900);

        // LOCK: prevents duplicate rows if this endpoint gets called twice
        // for the same payment in quick succession.
        const lock = LockService.getScriptLock();
        let result;
        try {
          lock.waitLock(15000);
          result = updateLeadStatus(p.phone || p.email, p.paymentId, paidAmount, isRecording);
          if (result === 'not_found') {
            saveLead(p.name, p.email, p.phone || '', p.paymentId, paidAmount, isRecording);
            result = 'updated';
          }
        } finally {
          lock.releaseLock();
        }
        // Only email/CAPI when this call actually flipped the row — avoids a
        // duplicate email if the Razorpay webhook already marked it Paid first.
        if (result === 'updated') {
          sendEmail(p.name, p.email, p.paymentId, isRecording);
          sendCAPIPurchase(String(p.paymentId).replace(/^RAZORPAY_/, ''), p.email, p.phone, isRecording);
        }
      } else {
        // Just initiated — save as Initiated (not Paid)
        saveLead(p.name, p.email, p.phone || '', 'INITIATED', null);
      }
    }
  } catch(err) {
    Logger.log('doGet error: ' + err.message);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'OK' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// 1. SAVE LEAD TO GOOGLE SHEET
// ------------------------------------------------------------
function saveLead(name, email, phone, paymentId, amount, isRecording) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const batchName = getActiveBatch();
  let sheet = ss.getSheetByName(batchName);

  if (!sheet) {
    sheet = ss.insertSheet(batchName);
    sheet.appendRow([
      '📅 Date & Time', '👤 Name', '📧 Email',
      '📱 WhatsApp', '💳 Payment ID', '💰 Amount',
      '✅ Status', '📝 Notes'
    ]);
    const header = sheet.getRange(1, 1, 1, 8);
    header.setBackground('#0d1b4c');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
    header.setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180); sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 220); sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 220); sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 120); sheet.setColumnWidth(8, 200);
  }

  const isPaid = paymentId && paymentId !== 'INITIATED' && paymentId !== 'PAYMENT_INITIATED' && paymentId !== 'LEAD';
  const status = isPaid ? (isRecording ? '✅ Paid ₹199 🎥' : '✅ Paid ₹99') : '🔄 Initiated';
  const bgColor = isPaid ? (isRecording ? '#e3f2fd' : '#e8f5e9') : '#fff9c4';

  const row = [
    new Date(),
    name,
    email,
    phone,
    paymentId || 'INITIATED',
    amount ? '₹' + (amount / 100) : (isPaid ? (isRecording ? '₹199' : '₹99') : '—'),
    status,
    isPaid ? (isRecording ? 'Paid ₹199 — Live + Recording 🎥' : 'Payment confirmed ✅') : 'Form filled — awaiting payment'
  ];
  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 8).setBackground(bgColor);
}

// ------------------------------------------------------------
// 2. UPDATE EXISTING ROW FROM INITIATED → PAID
// ------------------------------------------------------------
// Returns 'updated' | 'already_paid' | 'not_found'
function updateLeadStatus(phoneOrEmail, paymentId, amount, isRecording) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(getActiveBatch());
  if (!sheet) return 'not_found';

  const digitsOnly = String(phoneOrEmail).replace(/\D/g, '');
  const last10 = digitsOnly.slice(-10);

  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowPhone = String(data[i][3]).replace(/\D/g, '');
    const rowEmail = String(data[i][2]);
    const rowStatus = String(data[i][6]);

    const phoneMatch = !!last10 && rowPhone.slice(-10) === last10;
    const emailMatch = !!phoneOrEmail && rowEmail === phoneOrEmail;

    if (phoneMatch || emailMatch) {
      if (rowStatus.includes('Paid')) return 'already_paid';
      if (rowStatus.includes('Initiated')) {
        const isRec    = isRecording || amount >= 19900;
        const amtLabel  = isRec ? '₹199' : '₹99';
        const statusLbl = isRec ? '✅ Paid ₹199 🎥' : '✅ Paid ₹99';
        const notesLbl  = isRec ? 'Paid ₹199 — Live + Recording 🎥' : 'Payment confirmed ✅';
        const bgClr     = isRec ? '#e3f2fd' : '#e8f5e9';
        sheet.getRange(i + 1, 5).setValue(paymentId);
        sheet.getRange(i + 1, 6).setValue(amtLabel);
        sheet.getRange(i + 1, 7).setValue(statusLbl);
        sheet.getRange(i + 1, 8).setValue(notesLbl);
        sheet.getRange(i + 1, 1, 1, 8).setBackground(bgClr);
        return 'updated';
      }
    }
  }
  return 'not_found';
}

// ------------------------------------------------------------
// 3. SEND CONFIRMATION EMAIL
// ------------------------------------------------------------
function sendEmail(name, email, paymentId, isRecording) {
  const amountPaid = isRecording ? '₹199' : '₹99';
  const subject = '🎬 Payment Confirmed — Your AI Moviemaking Workshop Seat is Secured!';

  const htmlBody = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8"/>
    <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
    <style>
      body{margin:0;padding:0;background:#f4f4f4;font-family:'Helvetica Neue',Arial,sans-serif;}
      .wrap{max-width:560px;margin:32px auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08);}
      .header{background:linear-gradient(135deg,#0d1b4c,#1B4DB3);padding:36px 32px;text-align:center;}
      .header h1{color:#fff;font-size:26px;font-weight:900;margin:0 0 6px;}
      .header p{color:rgba(255,255,255,.75);font-size:14px;margin:0;}
      .tick{font-size:56px;margin-bottom:12px;}
      .body{padding:32px;}
      .greeting{font-size:20px;font-weight:800;color:#0d1b4c;margin-bottom:8px;}
      .msg{font-size:15px;color:#555;line-height:1.7;margin-bottom:24px;}
      .badge{background:#eaf2ff;border:2px solid #2F6FEB;border-radius:10px;padding:16px 20px;margin-bottom:24px;text-align:center;}
      .badge .amount{font-size:28px;font-weight:900;color:#0d1b4c;}
      .badge .label{font-size:13px;color:#777;margin-top:4px;}
      .steps{background:#f8f9ff;border-radius:12px;padding:20px;margin-bottom:24px;}
      .steps h3{font-size:13px;font-weight:800;color:#0d1b4c;text-transform:uppercase;letter-spacing:1px;margin:0 0 14px;}
      .step{display:flex;align-items:flex-start;gap:12px;margin-bottom:12px;}
      .step:last-child{margin-bottom:0;}
      .step-num{min-width:26px;height:26px;border-radius:50%;background:#0d1b4c;color:#fff;font-size:12px;font-weight:900;display:flex;align-items:center;justify-content:center;}
      .step-text{font-size:14px;color:#333;line-height:1.5;padding-top:3px;}
      .payment-id{background:#f5f5f5;border-radius:8px;padding:12px 16px;font-family:monospace;font-size:13px;color:#555;margin-bottom:24px;}
      .footer{background:#f8f8f8;border-top:1px solid #eee;padding:20px 32px;text-align:center;}
      .footer p{font-size:12px;color:#aaa;margin:0;line-height:1.6;}
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="header">
        <div class="tick">🎬</div>
        <h1>Payment Confirmed!</h1>
        <p>AI Moviemaking Workshop — Your seat is secured</p>
      </div>
      <div class="body">
        <div class="greeting">Hey ${name}! 👋</div>
        <p class="msg">
          You're officially in! Your payment has been received and your seat for the
          <strong>AI Moviemaking Workshop</strong> is confirmed.
          <br/><br/>
          In this workshop, you'll build your first cinematic AI film in 1 hour —
          no camera, no crew, no editing experience needed.
        </p>
        <div class="badge">
          <div class="amount">${amountPaid} Paid ✅</div>
          <div class="label">Payment ID: ${paymentId || 'CONFIRMED'}</div>
        </div>
        <!-- WhatsApp CTA -->
        <div style="background:linear-gradient(135deg,#e8f5e9,#d0f0da);border:2.5px solid #25D366;border-radius:16px;padding:20px;text-align:center;margin-bottom:24px;">
          <div style="font-size:13px;font-weight:800;color:#1a6b35;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">⚡ Step 1 — Join WhatsApp Group NOW</div>
          <div style="font-size:13px;color:#2e7d32;margin-bottom:14px;">Get the Zoom link, updates & reminders — all in the group</div>
          <a href="${CONFIG.WHATSAPP_GROUP}" style="display:block;background:#25D366;color:#fff;text-decoration:none;padding:16px 24px;border-radius:12px;font-size:16px;font-weight:900;box-shadow:0 6px 24px rgba(37,211,102,0.4);">💬 Join WhatsApp Group →</a>
        </div>

        <div class="steps">
          <h3>📋 What Happens Next</h3>
          <div class="step">
            <div class="step-num">1</div>
            <div class="step-text"><strong>Join the WhatsApp group</strong> above — tap the green button right now!</div>
          </div>
          <div class="step">
            <div class="step-num">2</div>
            <div class="step-text"><strong>Workshop date:</strong> ${CONFIG.WORKSHOP_DATE} — Live on Zoom.</div>
          </div>
          <div class="step">
            <div class="step-num">3</div>
            <div class="step-text"><strong>Show up and build your first AI film</strong> in 1 hour — no experience needed! 🚀</div>
          </div>
        </div>
        <p style="font-size:14px;color:#777;margin-bottom:8px;">Your Payment Reference:</p>
        <div class="payment-id">Payment ID: ${paymentId || 'CONFIRMED'}</div>
      </div>
      <div class="footer">
        <p>
          © AI Moviemaking Workshop by Alex<br/>
          110% Money Back Guarantee if you can't create your film in 1 hour.
        </p>
      </div>
    </div>
  </body>
  </html>
  `;

  GmailApp.sendEmail(email, subject, '', {
    htmlBody: htmlBody,
    name: CONFIG.SENDER_NAME,
    replyTo: Session.getActiveUser().getEmail()
  });
}
