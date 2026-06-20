// ============================================================
// Google Apps Script — Co-Working Booking Web App API
// Deploy as: Web App > Execute as: Me > Who has access: Anyone
// ============================================================

const SHEET_NAME = 'Bookings';
const BRANCH_SHEET_NAME = 'Branches';
const HEADERS = [
  'id', 'name', 'team', 'email', 'date', 'endDate', 'slot', 'slotLabel',
  'branch', 'branchId', 'user', 'password', 'status', 'timestamp', 'sendmail', 'cancelRequest'
];
const BRANCH_HEADERS = ['id', 'name', 'zone', 'address', 'travel', 'maps'];

const SPREADSHEET_ID = '1pSa5IxN3pjeltY35FOhw-pmttCr29qFzTAzHPz3Dees';

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0066FF').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  } else {
    // Sync headers — add any missing columns
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    HEADERS.forEach((h, idx) => {
      if (!existing.includes(h)) {
        const col = existing.length + 1;
        sheet.getRange(1, col).setValue(h).setFontWeight('bold').setBackground('#0066FF').setFontColor('#FFFFFF');
        existing.push(h);
      }
    });
  }
  return sheet;
}

function getBranchSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(BRANCH_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BRANCH_SHEET_NAME);
    sheet.appendRow(BRANCH_HEADERS);
    sheet.getRange(1, 1, 1, BRANCH_HEADERS.length).setFontWeight('bold').setBackground('#FF5B00').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllBranches() {
  const sheet = getBranchSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
    return obj;
  });
}

function seedBranches(branches) {
  const sheet = getBranchSheet();
  sheet.clearContents();
  sheet.appendRow(BRANCH_HEADERS);
  sheet.getRange(1, 1, 1, BRANCH_HEADERS.length).setFontWeight('bold').setBackground('#FF5B00').setFontColor('#FFFFFF');
  branches.forEach(b => sheet.appendRow(BRANCH_HEADERS.map(h => b[h] ?? '')));
  return { ok: true, count: branches.length };
}

function getAllBookings() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
    if (!obj.endDate) obj.endDate = obj.date;
    if (obj.cancelRequest === 'Confirm') obj.status = 'cancelled';
    return obj;
  });
}

// ── Email helper ─────────────────────────────────────────────
function formatDateTH(dateStr) {
  if (!dateStr) return '';
  const [y, m, d] = String(dateStr).split('-');
  const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                  'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return `${parseInt(d)} ${months[parseInt(m)-1]} ${parseInt(y)+543}`;
}

function sendBookingEmail(booking) {
  try {
    if (!booking.email) return;
    const isWaitlist = booking.status === 'waitlist';
    const endDate = booking.endDate && booking.endDate !== booking.date ? booking.endDate : null;
    const dateDisplay = endDate
      ? `${formatDateTH(booking.date)} – ${formatDateTH(endDate)}`
      : formatDateTH(booking.date);
    const days = endDate
      ? Math.ceil((new Date(endDate) - new Date(booking.date)) / 86400000) + 1
      : 1;

    const statusColor = isWaitlist ? '#D96500' : '#007A3E';
    const statusBg    = isWaitlist ? '#FFF4E6' : '#E6FDF4';
    const statusLabel = isWaitlist ? '⏳ คิวสำรอง' : '✅ ยืนยันแล้ว';
    const headerBg    = isWaitlist ? 'linear-gradient(135deg,#CF6B00,#FF8F00)' : 'linear-gradient(135deg,#0046CC,#0066FF)';

    const loginHtml = isWaitlist
      ? `<tr><td colspan="2" style="padding:16px;background:#FFF4E6;border-radius:8px;text-align:center;color:#D96500;font-size:13px;font-weight:600;">
           บัญชีเข้าใช้งานจะถูกปลดล็อคเมื่อได้รับการเลื่อนขั้นสิทธิ์เป็นคิวจริงค่ะ
         </td></tr>`
      : `<tr><td style="padding:8px 12px;color:#555;font-size:13px;width:120px;">Username</td>
           <td style="padding:8px 12px;font-weight:700;font-size:14px;font-family:monospace;">${booking.user || '-'}</td></tr>
         <tr><td style="padding:8px 12px;color:#555;font-size:13px;">Password</td>
           <td style="padding:8px 12px;font-weight:700;font-size:14px;font-family:monospace;">${booking.password || '-'}</td></tr>`;

    const daysRow = days > 1
      ? `<tr><td style="padding:8px 12px;color:#555;font-size:13px;">จำนวนวัน</td>
           <td style="padding:8px 12px;font-weight:700;font-size:14px;">${days} วัน</td></tr>`
      : '';

    const subject = isWaitlist
      ? `⏳ ลงทะเบียนคิวสำรอง Co-Working Space – ${booking.id}`
      : `✅ ยืนยันการจอง Co-Working Space – ${booking.id}`;

    const html = `
<!DOCTYPE html><html><head><meta charset="UTF-8"/></head><body style="margin:0;padding:0;background:#F5F7FA;font-family:'IBM Plex Sans Thai',Sarabun,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#F5F7FA;padding:32px 16px;">
<tr><td align="center">
<table width="560" cellpadding="0" cellspacing="0" style="max-width:560px;width:100%;">

  <!-- Header -->
  <tr><td style="background:${headerBg};border-radius:16px 16px 0 0;padding:32px 32px 24px;text-align:center;">
    <div style="font-size:11px;font-weight:700;letter-spacing:2px;color:rgba(255,255,255,0.7);text-transform:uppercase;margin-bottom:8px;">CO-WORKSPACE BOOKING</div>
    <div style="font-size:28px;font-weight:900;color:#fff;letter-spacing:1px;">${booking.id}</div>
    <div style="display:inline-block;margin-top:12px;background:rgba(255,255,255,0.2);border:1px solid rgba(255,255,255,0.3);border-radius:20px;padding:4px 16px;font-size:12px;font-weight:700;color:#fff;">${statusLabel}</div>
  </td></tr>

  <!-- Body -->
  <tr><td style="background:#fff;padding:32px;border-left:1px solid #E2E8F0;border-right:1px solid #E2E8F0;">

    <p style="margin:0 0 24px;font-size:15px;color:#334155;">สวัสดีคุณ <strong>${booking.name}</strong> 👋</p>
    <p style="margin:0 0 24px;font-size:14px;color:#64748B;line-height:1.6;">
      ${isWaitlist ? 'ระบบได้บันทึกรายชื่อท่านไว้ในคิวสำรองเรียบร้อยแล้วค่ะ เมื่อมีผู้สละสิทธิ์ ระบบจะปรับสถานะให้ท่านโดยอัตโนมัติค่ะ'
                  : 'ขอบคุณสำหรับการจองพื้นที่ Co-Working Space ค่ะ ด้านล่างคือรายละเอียดการจองของท่าน กรุณาบันทึกข้อมูลนี้ไว้ค่ะ'}
    </p>

    <!-- Booking details -->
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFC;border-radius:12px;border:1px solid #E2E8F0;margin-bottom:24px;overflow:hidden;">
      <tr><td colspan="2" style="padding:12px 16px;background:#F1F5F9;border-bottom:1px solid #E2E8F0;font-size:11px;font-weight:700;color:#64748B;letter-spacing:1px;text-transform:uppercase;">รายละเอียดการจอง</td></tr>
      <tr><td style="padding:8px 12px;color:#555;font-size:13px;width:120px;">ผู้จอง</td>
          <td style="padding:8px 12px;font-weight:700;font-size:14px;">${booking.name} ${booking.team ? `<span style="color:#94A3B8;font-size:12px;">(${booking.team})</span>` : ''}</td></tr>
      <tr style="background:#fff;"><td style="padding:8px 12px;color:#555;font-size:13px;">วันที่</td>
          <td style="padding:8px 12px;font-weight:700;font-size:14px;color:#0066FF;">${dateDisplay}</td></tr>
      ${daysRow}
      <tr><td style="padding:8px 12px;color:#555;font-size:13px;">ช่วงเวลา</td>
          <td style="padding:8px 12px;font-weight:700;font-size:14px;">${booking.slotLabel || '-'}</td></tr>
      <tr style="background:#fff;"><td style="padding:8px 12px;color:#555;font-size:13px;">สาขา</td>
          <td style="padding:8px 12px;font-weight:700;font-size:14px;">${booking.branch || '-'}</td></tr>
      <tr><td style="padding:8px 12px;color:#555;font-size:13px;">สถานะ</td>
          <td style="padding:8px 12px;"><span style="background:${statusBg};color:${statusColor};padding:3px 10px;border-radius:12px;font-size:12px;font-weight:700;">${statusLabel}</span></td></tr>
    </table>

    <!-- Login credentials -->
    ${!isWaitlist ? `
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#EFF6FF;border-radius:12px;border:1px solid #BFDBFE;margin-bottom:24px;overflow:hidden;">
      <tr><td colspan="2" style="padding:12px 16px;background:#DBEAFE;border-bottom:1px solid #BFDBFE;font-size:11px;font-weight:700;color:#1D4ED8;letter-spacing:1px;text-transform:uppercase;">สิทธิ์เข้าใช้งาน Wi-Fi & อาคาร</td></tr>
      ${loginHtml}
    </table>
    <p style="margin:0 0 24px;padding:12px 16px;background:#FFFBEB;border:1px solid #FDE68A;border-radius:10px;font-size:12px;color:#92400E;font-weight:600;">
      ⚠️ กรุณาเก็บรักษาข้อมูล Username / Password นี้ไว้เป็นความลับ และใช้เฉพาะในวันที่จองเท่านั้นค่ะ
    </p>` : `<table width="100%" cellpadding="0" cellspacing="0" style="background:#FFF4E6;border-radius:12px;border:1px solid #FFD8A8;margin-bottom:24px;overflow:hidden;">
      <tr><td style="padding:16px;font-size:13px;color:#D96500;">
        <strong>📌 หมายเหตุสำหรับคิวสำรอง:</strong><br/>
        ท่านอยู่ในรายการคิวสำรอง เมื่อมีผู้สละสิทธิ์ ระบบจะแจ้งเตือนท่านอีกครั้งทางอีเมลนี้ค่ะ
      </td></tr>
    </table>`}

    <!-- Reminder -->
    <table width="100%" cellpadding="0" cellspacing="0" style="background:#F0FDF4;border-radius:12px;border:1px solid #BBF7D0;margin-bottom:8px;overflow:hidden;">
      <tr><td style="padding:16px;font-size:12px;color:#166534;line-height:1.7;">
        📅 <strong>จองล่วงหน้าได้สูงสุด 14 วัน</strong><br/>
        ⚠️ <strong>การยกเลิก</strong>ต้องดำเนินการล่วงหน้าอย่างน้อย 2 วันก่อนวันเข้าใช้ค่ะ<br/>
        🗺️ ครอบคลุมพื้นที่ทั่วกรุงเทพและปริมณฑล 51 สาขา
      </td></tr>
    </table>
  </td></tr>

  <!-- Footer -->
  <tr><td style="background:#1E293B;border-radius:0 0 16px 16px;padding:20px 32px;text-align:center;">
    <div style="font-size:12px;color:rgba(255,255,255,0.5);">Co-Workspace Booking Smart Portal</div>
    <div style="font-size:11px;color:rgba(255,255,255,0.3);margin-top:4px;">อีเมลนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับค่ะ</div>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>`;

    MailApp.sendEmail({ to: booking.email, subject, htmlBody: html });
  } catch (e) {
    Logger.log('Email send error: ' + e.message);
  }
}

// ── REST handlers ────────────────────────────────────────────
function doGet(e) {
  const action = e?.parameter?.action || 'getAll';
  let result;
  try {
    if (action === 'getAll') {
      result = { ok: true, bookings: getAllBookings() };
    } else if (action === 'getBranches') {
      result = { ok: true, branches: getAllBranches() };
    } else {
      result = { ok: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch {
    return jsonResponse({ ok: false, error: 'Invalid JSON' });
  }

  const { action } = body;
  try {
    if (action === 'create') {
      return jsonResponse(createBooking(body.booking));
    } else if (action === 'cancel') {
      return jsonResponse(cancelBooking(body.id));
    } else if (action === 'requestCancel') {
      return jsonResponse(requestCancel(body.id));
    } else if (action === 'seedBranches') {
      return jsonResponse(seedBranches(body.branches));
    } else if (action === 'migrateBookingSheet') {
      return jsonResponse(migrateBookingSheet());
    } else if (action === 'seedCoPass') {
      return jsonResponse(seedCoPass());
    } else {
      return jsonResponse({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function getNextCoPassCredential() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const coPassSheet = ss.getSheetByName('Co-pass');
  if (!coPassSheet) return { id: '', password: '' };

  const coPassRows = coPassSheet.getDataRange().getValues().slice(1); // skip header
  if (!coPassRows.length) return { id: '', password: '' };

  // Count total confirmed bookings already saved to determine next index
  const bookingSheet = getSheet();
  const bookingData = bookingSheet.getDataRange().getValues();
  const bookingHeaders = bookingData[0];
  const statusCol = bookingHeaders.indexOf('status');
  const confirmedCount = bookingData.slice(1).filter(r => r[statusCol] === 'confirmed').length;

  const index = confirmedCount % coPassRows.length;
  return { id: coPassRows[index][0], password: coPassRows[index][1] };
}

function createBooking(booking) {
  const sheet = getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Assign Co-pass credentials for confirmed bookings
  let finalBooking = { ...booking };
  if (booking.status === 'confirmed') {
    const cred = getNextCoPassCredential();
    finalBooking.user = cred.id;
    finalBooking.password = cred.password;
  }

  const row = headers.map(h => finalBooking[h] ?? '');
  sheet.appendRow(row);
  sendBookingEmail(finalBooking);
  return { ok: true, booking: finalBooking };
}

function cancelBooking(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const statusCol = headers.indexOf('status');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      sheet.getRange(i + 1, statusCol + 1).setValue('cancelled');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Booking not found' };
}

function requestCancel(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const cancelRequestCol = headers.indexOf('cancelRequest');
  if (cancelRequestCol === -1) return { ok: false, error: 'cancelRequest column not found' };
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      sheet.getRange(i + 1, cancelRequestCol + 1).setValue('Requested');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Booking not found' };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Reorders Bookings sheet columns to match HEADERS exactly (non-destructive: missing values become '')
function migrateBookingSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: false, error: 'Bookings sheet not found' };

  const data    = sheet.getDataRange().getValues();
  const oldHdrs = data[0];
  const rows    = data.slice(1);

  // Map each row from old column order → new HEADERS order
  const newRows = rows.map(row => {
    return HEADERS.map(h => {
      const idx = oldHdrs.indexOf(h);
      return idx >= 0 ? row[idx] : '';
    });
  });

  sheet.clearContents();
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#0066FF')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, HEADERS.length).setValues(newRows);
  }

  return { ok: true, columns: HEADERS, rows: newRows.length };
}

function seedCoPass() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Co-pass');
  if (sheet) sheet.clearContents();
  else sheet = ss.insertSheet('Co-pass');

  const headers = ['id', 'password'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0066FF')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  const data = [
    ['User001', 'CW@Pass001'],
    ['User002', 'CW@Pass002'],
    ['User003', 'CW@Pass003'],
    ['User004', 'CW@Pass004'],
    ['User005', 'CW@Pass005'],
    ['User006', 'CW@Pass006'],
    ['User007', 'CW@Pass007'],
    ['User008', 'CW@Pass008'],
    ['User009', 'CW@Pass009'],
    ['User010', 'CW@Pass010'],
    ['User011', 'CW@Pass011'],
    ['User012', 'CW@Pass012'],
    ['User013', 'CW@Pass013'],
    ['User014', 'CW@Pass014'],
    ['User015', 'CW@Pass015'],
  ];
  sheet.getRange(2, 1, data.length, 2).setValues(data);

  return { ok: true, rows: data.length };
}
