const SHEETS = {
  X: '1E-oG-F6wRWmSFELnknvU_6wrkRfwhmnnQ8TCOhov1sg',
  XI: '1dtQlyFoUAGCPnlCY35owSU_08Mi7lV-7XRtXaiXKc7I',
  XII: '1jDG7OjWFt-eBLhhdlbRIrTCDgrw5Cn_HpU9Tj18JuKg',
  REKAP: '1Ho_WzEV3CASYK0sZ9TnmDjwBCU5s_ys8sOfNihqteM8'
};

const RESPONSE_HEADERS = [
  'nisn', 'nama', 'kelas', 'jk', 'tanggal',
  'bangunScore', 'ibadahScore', 'olahragaScore', 'makanScore',
  'belajarScore', 'sosialScore', 'tidurScore', 'totalPoin'
];

function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.action !== 'getData') {
      throw new Error('Gunakan action=getData.');
    }

    return jsonOutput({
      ok: true,
      students: readStudents(),
      submissions: readSubmissions()
    });
  } catch (error) {
    return jsonOutput({ ok: false, error: String(error.message || error) });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action !== 'submit') throw new Error('Action tidak valid.');

    const nisn = normalizeNisn(payload.nisn);
    if (!nisn) throw new Error('NISN wajib diisi.');
    if (!readStudents().some(student => student.nisn === nisn)) {
      throw new Error('NISN tidak ditemukan di spreadsheet data siswa.');
    }

    const sheet = getResponseSheet();
    const rowByHeader = {
      nisn: nisn,
      nama: String(payload.nama || ''),
      kelas: String(payload.kelas || ''),
      jk: String(payload.jk || ''),
      tanggal: String(payload.tanggal || ''),
      bangunscore: Number(payload.bangunScore) || 0,
      ibadahscore: Number(payload.ibadahScore) || 0,
      olahragascore: Number(payload.olahragaScore) || 0,
      makanscore: Number(payload.makanScore) || 0,
      belajarscore: Number(payload.belajarScore) || 0,
      sosialscore: Number(payload.sosialScore) || 0,
      tidurscore: Number(payload.tidurScore) || 0,
      totalpoin: Number(payload.totalPoin) || 0
    };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const row = headers.map(header => rowByHeader[compactHeader(header)] ?? '');

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }

    return jsonOutput({ ok: true });
  } catch (error) {
    return jsonOutput({ ok: false, error: String(error.message || error) });
  }
}

function readStudents() {
  return ['X', 'XI', 'XII'].flatMap(grade => {
    const sheet = SpreadsheetApp.openById(SHEETS[grade]).getSheets()[0];
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) return [];

    const headers = values[0].map(normalizeHeader);
    const nisnIndex = headers.findIndex(header => header.includes('nisn'));
    const nameIndex = headers.findIndex(header => header.includes('nama'));
    const classIndex = headers.findIndex(header => header.includes('kelas'));
    const genderIndex = headers.findIndex(header =>
      header === 'l/p' || header.includes('jenis kelamin') ||
      header.includes('gender') || header === 'jk'
    );
    if (nisnIndex < 0 || nameIndex < 0) {
      throw new Error(`Kolom NISN atau Nama tidak ditemukan di spreadsheet kelas ${grade}.`);
    }

    return values.slice(1).flatMap(row => {
      const nisn = normalizeNisn(row[nisnIndex]);
      if (!nisn) return [];
      return [{
        nisn: nisn,
        nama: String(row[nameIndex] || '').trim(),
        kelas: String((classIndex >= 0 && row[classIndex]) || grade).trim(),
        jk: String((genderIndex >= 0 && row[genderIndex]) || '').trim().toUpperCase().startsWith('L') ? 'L' : 'P'
      }];
    });
  });
}

function readSubmissions() {
  const sheet = getResponseSheet();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0].map(normalizeHeader);
  const nisnIndex = headers.findIndex(header => header.includes('nisn'));
  const scoreIndex = headers.findIndex(header =>
    header.includes('poin') || header.includes('skor') ||
    header.includes('nilai') || header.includes('total')
  );
  const dateIndex = headers.findIndex(header => header.includes('tanggal') || header.includes('date'));
  if (nisnIndex < 0) throw new Error('Kolom NISN tidak ditemukan di spreadsheet responden.');

  return values.slice(1).flatMap(row => {
    const nisn = normalizeNisn(row[nisnIndex]);
    if (!nisn) return [];
    return [{
      nisn: nisn,
      totalPoin: scoreIndex >= 0 ? Number(String(row[scoreIndex]).replace(/[^0-9.-]/g, '')) || 0 : 0,
      tanggal: dateIndex >= 0 ? String(row[dateIndex] || '').trim() : ''
    }];
  });
}

function getResponseSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEETS.REKAP);
  const sheet = spreadsheet.getSheets()[0];
  const lastColumn = sheet.getLastColumn();
  const existingHeaders = lastColumn
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(compactHeader)
    : [];
  const headers = existingHeaders.filter(Boolean);

  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
  } else {
    const missing = RESPONSE_HEADERS.filter(header => !headers.includes(compactHeader(header)));
    if (missing.length) {
      sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
    }
  }
  return sheet;
}

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase();
}

function compactHeader(value) {
  return normalizeHeader(value).replace(/[\s_-]/g, '');
}

function normalizeNisn(value) {
  const digits = String(value || '').replace(/\D/g, '');
  return digits ? digits.padStart(10, '0') : '';
}

function jsonOutput(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
