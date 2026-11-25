// Jadwalin - server. Tempel di Google Sheets: Ekstensi > Apps Script. Lalu Deploy > Web app.
// Sheet "Users", "Sesi", dan "Konten" dibuat/diperbarui otomatis saat pertama dipakai.
const HEAD = {
  Users: ["email", "nama", "salt", "hash", "dibuat", "gagal", "kunci_sampai", "kode", "kode_sampai", "foto", "sosmed", "target"],
  Sesi: ["token", "email", "kedaluwarsa"],
  Konten: ["email", "id", "judul", "platform", "tahap", "deadline", "catatan", "sow", "kolab", "fee", "mulai", "kirim", "brand", "pic", "kontak", "bayar", "dibayar", "keranjang", "boost", "boost_hari", "link_produk", "link_brief", "keranjang_putih", "taplink"],
  Pengeluaran: ["email", "id", "tanggal", "nama", "jumlah"]
};
const ready_ = {};

// PENTING: jalankan fungsi ini SEKALI dari editor (pilih "izinkanEmail" > Run) agar Google
// meminta izin mengirim email. Tanpa ini, fitur lupa kata sandi akan gagal mengirim email.
function izinkanEmail() {
  const me = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail(me, "Tes email Jadwalin", "Izin kirim email sudah aktif. Fitur lupa kata sandi siap dipakai.");
  Logger.log("Berhasil. Email tes dikirim ke " + me + ". Sisa kuota harian: " + MailApp.getRemainingDailyQuota());
}

function sheet_(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const w = HEAD[n].length;
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    if (s.getMaxColumns() < w) s.insertColumnsAfter(s.getMaxColumns(), w - s.getMaxColumns());
    s.getRange(1, 1, s.getMaxRows(), w).setNumberFormat("@"); // semua teks, cegah rumus
    s.getRange(1, 1, 1, w).setValues([HEAD[n]]);
    s.setFrozenRows(1);
    ready_[n] = true;
    return s;
  }
  if (!ready_[n]) { // sheet lama: tambahkan kolom baru otomatis
    if (s.getMaxColumns() < w) s.insertColumnsAfter(s.getMaxColumns(), w - s.getMaxColumns());
    const h = s.getRange(1, 1, 1, w);
    if (h.getValues()[0].join("|") !== HEAD[n].join("|")) {
      s.getRange(1, 1, s.getMaxRows(), w).setNumberFormat("@");
      h.setValues([HEAD[n]]);
    }
    ready_[n] = true;
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function c_(v, n) { return String(v == null ? "" : v).substring(0, n); }
function s_(v, n) { return c_(v, n).replace(/[\u0000-\u001f]/g, " ").trim(); }
// Satu link: hanya http(s) atau alamat tanpa skema (mis. tokopedia.link/abc)
function link_(v) {
  const x = s_(v, 300);
  if (!x) return "";
  if (/^[a-z][a-z0-9+.\-]*:/i.test(x) && !/^https?:\/\//i.test(x)) return "";
  return x;
}
// Beberapa link, satu per baris (maksimal 10)
function links_(v) {
  return String(v == null ? "" : v).split(/\r?\n/).map(x => link_(x)).filter(Boolean).slice(0, 10).join("\n").substring(0, 3000);
}
function norm_(s) { return String(s || "").trim().toLowerCase(); }
function hash_(pw, salt) {
  let h = salt + pw;
  for (let i = 0; i < 1000; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + h));
  }
  return h;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (x) { return out_({ ok: false, error: "Server sibuk, coba lagi." }); }
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "register": return out_(register_(d));
      case "login": return out_(login_(d));
      case "load": return out_(load_(d));
      case "save": return out_(save_(d));
      case "logout": return out_(logout_(d));
      case "forgot": return out_(forgot_(d));
      case "reset": return out_(reset_(d));
      case "profile": return out_(profile_(d));
      case "changepw": return out_(changepw_(d));
    }
    return out_({ ok: false, error: "Aksi tidak dikenal" });
  } catch (err) {
    console.error(err);
    return out_({ ok: false, error: "Terjadi kesalahan" });
  } finally {
    lock.releaseLock();
  }
}

function findUser_(email) {
  const s = sheet_("Users"), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === email) return { row: i + 1, v: v[i], s: s };
  return null;
}
function newSession_(email) {
  const t = Utilities.getUuid() + Utilities.getUuid();
  sheet_("Sesi").appendRow([t, email, String(Date.now() + 30 * 864e5)]);
  return t;
}
function sessionEmail_(token) {
  const s = sheet_("Sesi"), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === token) {
      if (Number(v[i][2]) > Date.now()) return v[i][1];
      s.deleteRow(i + 1);
      return null;
    }
  }
  return null;
}
function tasks_(email) {
  const v = sheet_("Konten").getDataRange().getValues().slice(1);
  return v.filter(r => r[0] === email).map(r => ({
    id: r[1], title: r[2], platform: r[3], stage: r[4], due: r[5], note: r[6], sow: r[7],
    kolab: r[8] || "", fee: r[9] || "", start: r[10] || "", kirim: r[11] || "",
    brand: r[12] || "", pic: r[13] || "", kontak: r[14] || "", bayar: r[15] || "", dibayar: r[16] || "",
    keranjang: r[17] || "", boost: r[18] || "", boostDays: r[19] || "", linkProduk: r[20] || "", linkBrief: r[21] || "",
    keranjangPutih: r[22] || "", taplink: r[23] || ""
  }));
}
function expenses_(email) {
  const v = sheet_("Pengeluaran").getDataRange().getValues().slice(1);
  return v.filter(r => r[0] === email).map(r => ({ id: r[1], date: r[2], name: r[3], amount: r[4] || "" }));
}
// Paket data yang dikirim ke web setelah masuk
function pack_(email, v, token) {
  let links = {};
  try { links = JSON.parse(v[10] || "{}"); } catch (e) {}
  return { ok: true, token: token, email: email, name: v[1], photo: String(v[9] || ""), links: links, tasks: tasks_(email), expenses: expenses_(email), target: String(v[11] || "") };
}

function register_(d) {
  const email = norm_(d.email), pw = String(d.password || "");
  if (!/^[a-z0-9][^\s@]*@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: "Email tidak valid" };
  if (pw.length < 6) return { ok: false, error: "Kata sandi minimal 6 karakter" };
  if (findUser_(email)) return { ok: false, error: "Email sudah terdaftar" };
  const name = c_(d.name, 60).trim() || email.split("@")[0];
  const salt = Utilities.getUuid();
  sheet_("Users").appendRow([email, name, salt, hash_(pw, salt), new Date().toISOString(), "0", "0"]);
  return { ok: true, token: newSession_(email), email: email, name: name, photo: "", links: {}, tasks: [], expenses: [], target: "" };
}
function login_(d) {
  const email = norm_(d.email), u = findUser_(email);
  const bad = { ok: false, error: "Email atau kata sandi salah" };
  if (!u) return bad;
  if (Number(u.v[6]) > Date.now()) return { ok: false, error: "Terlalu banyak percobaan. Coba lagi 10 menit lagi." };
  if (hash_(String(d.password || ""), u.v[2]) !== u.v[3]) {
    let f = Number(u.v[5]) + 1, k = 0;
    if (f >= 5) { k = Date.now() + 600000; f = 0; }
    u.s.getRange(u.row, 6, 1, 2).setValues([[String(f), String(k)]]);
    return bad;
  }
  u.s.getRange(u.row, 6, 1, 2).setValues([["0", "0"]]);
  return pack_(email, u.v, newSession_(email));
}
function load_(d) {
  const email = sessionEmail_(d.token);
  if (!email) return { ok: false, expired: true, error: "Sesi berakhir" };
  const u = findUser_(email);
  if (!u) return { ok: false, expired: true, error: "Akun tidak ditemukan" };
  return pack_(email, u.v, d.token);
}
function save_(d) {
  const email = sessionEmail_(d.token);
  if (!email) return { ok: false, expired: true, error: "Sesi berakhir" };
  const W = HEAD.Konten.length;
  const s = sheet_("Konten"), v = s.getDataRange().getValues().slice(1);
  const keep = v.filter(r => r[0] !== email).map(r => { const x = r.slice(0, W); while (x.length < W) x.push(""); return x; });
  const add = (Array.isArray(d.tasks) ? d.tasks : []).slice(0, 500).map(t => {
    const paid = t.kolab === "paid", barter = t.kolab === "barter", tiktok = t.platform === "TikTok";
    const igFeed = ["Instagram Carousel", "Instagram Reels", "Instagram", "Reels"].indexOf(t.platform) >= 0;
    const igStory = t.platform === "Instagram Story";
    const bayar = paid ? (["belum", "dp", "lunas"].indexOf(t.bayar) >= 0 ? t.bayar : "belum") : "";
    return [
      email, c_(t.id, 40), c_(t.title, 120), c_(t.platform, 30), c_(t.stage, 20), c_(t.due, 10), c_(t.note, 300), c_(t.sow, 1500),
      paid || barter ? t.kolab : "",
      paid ? String(t.fee || "").replace(/\D/g, "").substring(0, 12) : "",
      /^\d{4}-\d{2}-\d{2}$/.test(String(t.start || "")) ? t.start : "",
      barter && ["belum", "jalan", "sampai"].indexOf(t.kirim) >= 0 ? t.kirim : (barter ? "belum" : ""),
      paid || barter ? s_(t.brand, 60) : "",
      paid || barter ? s_(t.pic, 60) : "",
      paid || barter ? s_(t.kontak, 80) : "",
      bayar,
      bayar === "dp" ? String(t.dibayar || "").replace(/\D/g, "").substring(0, 12) : "",
      tiktok && ["ya", "belum", "pasang"].indexOf(t.keranjang) >= 0 ? "ya" : "", // data lama (belum/pasang) jadi "ya"
      tiktok ? s_(t.boost, 60) : "",
      tiktok ? String(t.boostDays || "").replace(/\D/g, "").substring(0, 3) : "",
      links_(t.linkProduk || t.link),
      link_(t.linkBrief),
      igFeed && t.keranjangPutih === "ya" ? "ya" : "",
      igStory && t.taplink === "ya" ? "ya" : ""
    ];
  });
  const all = keep.concat(add);
  if (s.getMaxRows() > 1) s.getRange(2, 1, s.getMaxRows() - 1, W).clearContent();
  if (all.length) {
    if (s.getMaxRows() < all.length + 1) s.insertRowsAfter(s.getMaxRows(), all.length + 1 - s.getMaxRows());
    const r = s.getRange(2, 1, all.length, W);
    r.setNumberFormat("@");
    r.setValues(all);
  }
  // Pengeluaran: hanya diganti kalau web mengirimnya (web lama tidak menghapus data)
  if (Array.isArray(d.expenses)) {
    const es = sheet_("Pengeluaran"), EW = HEAD.Pengeluaran.length;
    const ev = es.getDataRange().getValues().slice(1);
    const ekeep = ev.filter(r => r[0] !== email).map(r => r.slice(0, EW));
    const eadd = d.expenses.slice(0, 500).map(x => [
      email, c_(x.id, 40),
      /^\d{4}-\d{2}-\d{2}$/.test(String(x.date || "")) ? x.date : "",
      s_(x.name, 80),
      String(x.amount || "").replace(/\D/g, "").substring(0, 12)
    ]);
    const eall = ekeep.concat(eadd);
    if (es.getMaxRows() > 1) es.getRange(2, 1, es.getMaxRows() - 1, EW).clearContent();
    if (eall.length) {
      if (es.getMaxRows() < eall.length + 1) es.insertRowsAfter(es.getMaxRows(), eall.length + 1 - es.getMaxRows());
      const er = es.getRange(2, 1, eall.length, EW);
      er.setNumberFormat("@");
      er.setValues(eall);
    }
  }
  // Target profit bulanan
  if (d.target !== undefined) {
    const u = findUser_(email);
    if (u) {
      const tc = u.s.getRange(u.row, 12);
      tc.setNumberFormat("@");
      tc.setValue(String(d.target || "").replace(/\D/g, "").substring(0, 12));
    }
  }
  return { ok: true };
}
function logout_(d) {
  const s = sheet_("Sesi"), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === d.token) s.deleteRow(i + 1);
  return { ok: true };
}
function logoutAll_(email) {
  const s = sheet_("Sesi"), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][1] === email) s.deleteRow(i + 1);
}

// Simpan profil: nama, foto (data URL kecil), dan link sosmed
function profile_(d) {
  const email = sessionEmail_(d.token);
  if (!email) return { ok: false, expired: true, error: "Sesi berakhir" };
  const u = findUser_(email);
  if (!u) return { ok: false, expired: true, error: "Akun tidak ditemukan" };
  const name = c_(d.name, 60).trim() || String(u.v[1]);
  const photo = String(d.photo || "");
  if (photo && (photo.length > 45000 || !/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(photo))) {
    return { ok: false, error: "Foto terlalu besar atau formatnya tidak didukung" };
  }
  const L = d.links && typeof d.links === "object" ? d.links : {};
  const links = {};
  ["tiktok", "instagram", "threads", "youtube", "x", "lainnya"].forEach(k => {
    const val = c_(L[k], 150).replace(/[\u0000-\u001f]/g, "").trim();
    if (val) links[k] = val;
  });
  u.s.getRange(u.row, 2).setValue(name);
  const r = u.s.getRange(u.row, 10, 1, 2);
  r.setNumberFormat("@");
  r.setValues([[photo, JSON.stringify(links)]]);
  return { ok: true, name: name, photo: photo, links: links };
}
// Ganti kata sandi dari halaman profil (wajib kata sandi lama)
function changepw_(d) {
  const email = sessionEmail_(d.token);
  if (!email) return { ok: false, expired: true, error: "Sesi berakhir" };
  const u = findUser_(email);
  if (!u) return { ok: false, expired: true, error: "Akun tidak ditemukan" };
  const np = String(d.newPassword || "");
  if (np.length < 6) return { ok: false, error: "Kata sandi baru minimal 6 karakter" };
  if (Number(u.v[6]) > Date.now()) return { ok: false, error: "Terlalu banyak percobaan. Coba lagi 10 menit lagi." };
  if (hash_(String(d.oldPassword || ""), u.v[2]) !== u.v[3]) {
    let f = Number(u.v[5]) + 1, k = 0;
    if (f >= 5) { k = Date.now() + 600000; f = 0; }
    u.s.getRange(u.row, 6, 1, 2).setValues([[String(f), String(k)]]);
    return { ok: false, error: "Kata sandi lama salah" };
  }
  const salt = Utilities.getUuid();
  u.s.getRange(u.row, 3, 1, 2).setValues([[salt, hash_(np, salt)]]);
  u.s.getRange(u.row, 6, 1, 2).setValues([["0", "0"]]);
  logoutAll_(email); // perangkat lain otomatis keluar
  return { ok: true, token: newSession_(email) };
}

// Lupa kata sandi: kirim kode 6 angka ke email (berlaku 15 menit)
function forgot_(d) {
  const email = norm_(d.email), u = findUser_(email);
  const msg = "Jika email terdaftar, kode 6 angka sudah dikirim (berlaku 15 menit). Cek juga folder spam.";
  if (!u) return { ok: true, message: msg }; // sama persis, agar email terdaftar tidak bisa ditebak
  if (Number(u.v[8] || 0) - 14 * 60000 > Date.now()) return { ok: true, message: msg }; // baru dikirim <1 menit lalu
  const code = String(100000 + (parseInt(Utilities.getUuid().replace(/-/g, "").substring(0, 8), 16) % 900000));
  try {
    MailApp.sendEmail(email, "Kode reset kata sandi Jadwalin",
      "Kode untuk mengatur ulang kata sandimu: " + code + "\n\nKode berlaku 15 menit. Jika bukan kamu yang meminta, abaikan email ini.");
  } catch (err) {
    console.error("Gagal kirim email: " + err);
    return { ok: false, error: "Email gagal dikirim (" + String(err && err.message || err).substring(0, 140) + ")" };
  }
  const r = u.s.getRange(u.row, 8, 1, 2);
  r.setNumberFormat("@");
  r.setValues([[hash_(code, u.v[2]), String(Date.now() + 15 * 60000)]]);
  return { ok: true, message: msg };
}
function reset_(d) {
  const email = norm_(d.email), u = findUser_(email), pw = String(d.password || "");
  const bad = { ok: false, error: "Kode salah atau sudah kedaluwarsa" };
  if (pw.length < 6) return { ok: false, error: "Kata sandi minimal 6 karakter" };
  if (!u || !u.v[7] || Number(u.v[8]) < Date.now()) return bad;
  if (Number(u.v[6]) > Date.now()) return { ok: false, error: "Terlalu banyak percobaan. Coba lagi 10 menit lagi." };
  if (hash_(String(d.code || "").trim(), u.v[2]) !== u.v[7]) {
    let f = Number(u.v[5]) + 1, k = 0;
    if (f >= 5) { k = Date.now() + 600000; f = 0; }
    u.s.getRange(u.row, 6, 1, 2).setValues([[String(f), String(k)]]);
    if (k) u.s.getRange(u.row, 8, 1, 2).setValues([["", ""]]); // kode dihanguskan setelah 5x salah
    return bad;
  }
  const salt = Utilities.getUuid();
  u.s.getRange(u.row, 3, 1, 2).setValues([[salt, hash_(pw, salt)]]);
  u.s.getRange(u.row, 6, 1, 4).setValues([["0", "0", "", ""]]);
  logoutAll_(email); // semua sesi lama otomatis keluar
  return pack_(email, u.v, newSession_(email));
}
