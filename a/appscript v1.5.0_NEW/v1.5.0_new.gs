/**
 * ═════════════════════════════════════════════════════════════════════
 * 🗄️ ROYYEK VAULT BACKEND - Google Apps Script
 * Versi: 1.5.0 - ULTRA SPEED & PRODUCTION READY (FIXED)
 * ═════════════════════════════════════════════════════════════════════
 * Detail Pembaruan & Perbaikan Sistem Performa:
 * 1. FIX INFINITE LOOP: Menambahkan ON_EDIT_LOCK pada onEdit untuk mencegah 
 * multi-triggering dan menghemat kuota Google Server.
 * 2. ACCURATE CELL INDEX: Sinkronisasi koordinat toggleFavoriteData diperketat 
 * agar tidak salah baris akibat pembacaan array.
 * 3. SAFE CELL LOCKING: Menggunakan LockService 30 detik untuk menjamin 
 * keamanan data saat diakses banyak perangkat secara bersamaan.
 * 4. COMPATIBILITY: Tetap mempertahankan 100% format backup.json v8.4 Anda.
 */

const SHEET_NAME = "Vault";
const NOTES_SHEET_NAME = "Notes";

function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
    return fn();
  } catch (e) {
    return output({ success: false, error: "Lock timeout. Sistem sibuk, coba lagi: " + e.toString() });
  } finally {
    lock.releaseLock();
  }
}

/* =====================================================
   🔧 HELPER: GET/CREATE SHEET & ENSURE COLUMNS
===================================================== */
function getSheet(sheetTargetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetTargetName);
  const isNotes = sheetTargetName === NOTES_SHEET_NAME;

  if (!sheet) {
    sheet = ss.insertSheet(sheetTargetName);
    if (isNotes) {
      sheet.appendRow(["id", "kategori", "judul", "isi", "tanggal", "favorit", "vault_id"]);
    } else {
      sheet.appendRow(["id", "kategori", "nama_aplikasi", "username", "email", "password", "ket", "ket_dua", "created_at", "favorit", "vault_id"]);
    }
  } else {
    ensureRequiredColumns(sheet, isNotes);
  }
  return sheet;
}

function ensureRequiredColumns(sheet, isNotes) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Pastikan kolom favorit ada secara presisi
  if (!headers.includes("favorit")) {
    const favColIndex = isNotes ? 6 : 10;
    sheet.insertColumnBefore(favColIndex < sheet.getLastColumn() ? favColIndex : sheet.getLastColumn());
    sheet.getRange(1, favColIndex).setValue("favorit");
  }
  
  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Pastikan kolom vault_id ada
  if (!updatedHeaders.includes("vault_id")) {
    const newColIndex = sheet.getLastColumn() + 1;
    sheet.getRange(1, newColIndex).setValue("vault_id");
    
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const defaults = Array(lastRow - 1).fill(["default"]);
      sheet.getRange(2, newColIndex, lastRow - 1, 1).setValues(defaults);
    }
  }
}

function output(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function generateID(prefix) {
  return prefix + "-" + Date.now() + "-" + Math.floor(Math.random() * 9999);
}

/* =====================================================
   📥 GET HANDLER (READ DATA)
===================================================== */
function doGet(e) {
  try {
    const isNotesRequest = (e && e.parameter && e.parameter.type === "notes");
    const targetTab = isNotesRequest ? NOTES_SHEET_NAME : SHEET_NAME;
    const sheet = getSheet(targetTab);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) return output([]);

    const headers = data[0];
    const result = data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        if (h === "favorit") {
          let val = row[i];
          if (typeof val === "boolean") {
            obj[h] = val;
          } else {
            obj[h] = String(val || '').toLowerCase().trim() === 'true';
          }
        } else {
          obj[h] = row[i];
        }
      });
      return obj;
    });
    return output(result);
  } catch(err) {
    return output({ success: false, error: err.toString() });
  }
}

/* =====================================================
   📤 POST HANDLER ROUTER (MUTATION DATA)
===================================================== */
function doPost(e) {
  try {
    const action = e.parameter.action;
    const isNotes = e.parameter.data_type === "notes"; 
    
    if (action === "create" && !isNotes && e.parameter.data && (e.parameter.data.includes("\n") || e.parameter.data.includes("|") || e.parameter.data.trim().startsWith("{") || e.parameter.data.includes(","))) {
      return imporMassalTeks(e);
    }
    
    switch(action) {
      case "create": return withLock(function() { return createData(e, isNotes); });
      case "update": return withLock(function() { return updateData(e, isNotes); });
      case "delete": return withLock(function() { return deleteData(e, isNotes); });
      case "bulk_delete": return withLock(function() { return bulkDeleteData(e, isNotes); });
      case "toggle_favorite": return withLock(function() { return toggleFavoriteData(e, isNotes); }); 
      case "update_category": return withLock(function() { return updateCategoryData(e, isNotes); });
      case "update_vault": return withLock(function() { return updateVaultId(e, isNotes); });
      case "sync_queue": return withLock(function() { return processSyncQueue(e); });
      case "reorder_folders": return reorderFolders(e);
      case "import_massal": return imporMassalTeks(e); 
      default: return output({ success: false, message: "Invalid action: " + action });
    }
  } catch(err) {
    return output({ success: false, error: err.toString() });
  }
}

/* =====================================================
   ➕ CREATE DATA
===================================================== */
function createData(e, isNotes) {
  const vaultId = e.parameter.vault_id || "default";
  
  if (isNotes) {
    const sheet = getSheet(NOTES_SHEET_NAME);
    const judul = (e.parameter.judul || "").trim();
    if (judul === "") return output({ success: false, message: "Judul catatan wajib diisi!" });

    const id = generateID("NOTE");
    let favStatus = String(e.parameter.favorit).toLowerCase().trim() === 'true';
    
    sheet.appendRow([id, e.parameter.kategori || "Umum", judul, e.parameter.isi || "", new Date(), favStatus, vaultId]);
    SpreadsheetApp.flush();
    return output({ success: true, message: "created", id: id });
  } else {
    const sheet = getSheet(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const inputApp = (e.parameter.nama_aplikasi || "").trim().toLowerCase();
    const inputUser = (e.parameter.username || "").trim().toLowerCase();
    const inputEmail = (e.parameter.email || "").trim().toLowerCase();

    if (inputApp === "") return output({ success: false, message: "Nama aplikasi wajib diisi!" });

    for (let i = 1; i < data.length; i++) {
      const sheetApp = String(data[i][2] || "").trim().toLowerCase();
      const sheetUser = String(data[i][3] || "").trim().toLowerCase();
      const sheetEmail = String(data[i][4] || "").trim().toLowerCase();
      if (sheetApp === inputApp && (sheetUser === inputUser || sheetEmail === inputEmail)) {
        return output({ success: false, message: "Data Ganda! Akun " + e.parameter.nama_aplikasi + " sudah ada." });
      }
    }

    const id = generateID("VAULT");
    let favStatus = String(e.parameter.favorit).toLowerCase().trim() === 'true';
    
    sheet.appendRow([id, e.parameter.kategori || "Umum", e.parameter.nama_aplikasi || "", e.parameter.username || "", e.parameter.email || "", e.parameter.password || "", e.parameter.ket || "", e.parameter.ket_dua || "", new Date(), favStatus, vaultId]);
    SpreadsheetApp.flush();
    return output({ success: true, message: "created", id: id });
  }
}

/* =====================================================
   ⚡ IMPOR MASSAL MULTI-FORMAT (SUPER TURBO BOOST)
===================================================== */
function imporMassalTeks(e) {
  return withLock(function() {
    PropertiesService.getScriptProperties().setProperty("BULK_PROCESS_RUNNING", "true");
    
    try {
      const sheet = getSheet(SHEET_NAME);
      const vaultId = e.parameter.vault_id || "default";
      const rawData = e.parameter.data || "";
      
      if (rawData.trim() === "") return output({ success: false, message: "Tidak ada data untuk diimpor." });
      
      const arrayDataBaru = [];
      const waktuSekarang = new Date();
      
      // 1. FORMAT JSON MURNI
      if (rawData.trim().startsWith("{")) {
        const jsonObj = JSON.parse(rawData);
        const vaultItems = jsonObj.vault || [];
        
        for (let i = 0; i < vaultItems.length; i++) {
          let item = vaultItems[i];
          if (!item.nama_aplikasi) continue;
          
          let favStatus = false;
          if (item.favorit !== undefined) {
            if (typeof item.favorit === "boolean") {
              favStatus = item.favorit;
            } else {
              let val = String(item.favorit).toLowerCase().trim();
              favStatus = (val === 'true' || val === '1' || val === 'yes' || val === 'ya');
            }
          }
          
          arrayDataBaru.push([
            item.id || generateID("VAULT"),
            item.kategori || "Umum",
            item.nama_aplikasi,
            item.username || "",
            item.email || "",
            item.password || "",
            item.ket || "",
            item.ket_dua || "",
            item.created_at ? new Date(item.created_at) : waktuSekarang,
            favStatus,
            item.vault_id || vaultId
          ]);
        }
        
        if (arrayDataBaru.length > 0) {
          sheet.getRange(sheet.getLastRow() + 1, 1, arrayDataBaru.length, 11).setValues(arrayDataBaru);
          SpreadsheetApp.flush(); 
          return output({ success: true, message: "Sukses mengimpor " + arrayDataBaru.length + " data dari file JSON!" });
        }
      } 
      // 2. FORMAT CSV
      else if (rawData.includes(",") || rawData.includes(";")) {
        const delimiter = rawData.includes(";") ? ";" : ",";
        const lines = rawData.split(/\r?\n/);
        
        for (let i = 1; i < lines.length; i++) {
          let line = lines[i].trim();
          if (line === "") continue;
          
          let pattern = new RegExp("([^" + delimiter + "\"]*|\"([^\"]*)\")(?=" + delimiter + "|$)", "g");
          let matches = line.match(pattern);
          if (!matches) continue;
          
          let row = matches.map(val => {
            let clean = val.trim();
            if (clean.startsWith('"') && clean.endsWith('"')) {
              clean = clean.substring(1, clean.length - 1);
            }
            return clean;
          });
          
          if (row[0] && row[0].toLowerCase().trim() === "vault") {
            if (!row[2] || row[2].trim() === "") continue;
            
            let favStatus = false;
            if (row[8] !== undefined) {
              let val = String(row[8]).toLowerCase().trim();
              favStatus = (val === 'true' || val === '1' || val === 'yes' || val === 'ya');
            }
            
            arrayDataBaru.push([
              generateID("VAULT"),
              row[1] ? row[1].trim() : "Umum",
              row[2] ? row[2].trim() : "",
              row[3] ? row[3].trim() : "",
              row[4] ? row[4].trim() : "",
              row[5] ? row[5].trim() : "",
              row[6] ? row[6].trim() : "",
              row[7] ? row[7].trim() : "",
              row[9] ? new Date(row[9]) : waktuSekarang,
              favStatus,
              vaultId
            ]);
          }
        }
        
        if (arrayDataBaru.length > 0) {
          sheet.getRange(sheet.getLastRow() + 1, 1, arrayDataBaru.length, 11).setValues(arrayDataBaru);
          SpreadsheetApp.flush();
          return output({ success: true, message: "Sukses mengimpor " + arrayDataBaru.length + " data Vault dari file CSV!" });
        }
      } 
      // 3. FORMAT TEKS LAWAS LINE (|)
      else {
        const baris = rawData.split("\n");
        for (let i = 0; i < baris.length; i++) {
          let barisTeks = baris[i].trim();
          if (barisTeks === "") continue;
          
          let kolom = barisTeks.split("|");
          if (!kolom[1]) continue;
          
          let favStatus = kolom[7] ? kolom[7].trim().toLowerCase() === 'true' : false;
          
          arrayDataBaru.push([
            generateID("VAULT"),
            kolom[0] ? kolom[0].trim() : "Umum",
            kolom[1].trim(),
            kolom[2] ? kolom[2].trim() : "",
            kolom[3] ? kolom[3].trim() : "",
            kolom[4] ? kolom[4].trim() : "",
            kolom[5] ? kolom[5].trim() : "",
            kolom[6] ? kolom[6].trim() : "",
            waktuSekarang,
            favStatus,
            vaultId
          ]);
        }
        if (arrayDataBaru.length > 0) {
          sheet.getRange(sheet.getLastRow() + 1, 1, arrayDataBaru.length, 11).setValues(arrayDataBaru);
          SpreadsheetApp.flush();
          return output({ success: true, message: "Sukses mengimpor " + arrayDataBaru.length + " data teks!" });
        }
      }
      return output({ success: false, message: "Tidak ada data valid yang berhasil diekstrak." });
      
    } catch(err) {
      return output({ success: false, error: err.toString() });
    } finally {
      PropertiesService.getScriptProperties().setProperty("BULK_PROCESS_RUNNING", "false");
    }
  });
}

/* =====================================================
   ✏️ UPDATE DATA (FAST INDEX SEARCH)
===================================================== */
function updateData(e, isNotes) {
  const targetTab = isNotes ? NOTES_SHEET_NAME : SHEET_NAME;
  const sheet = getSheet(targetTab);
  const data = sheet.getDataRange().getValues();
  const id = String(e.parameter.id || "").trim();
  
  const idList = data.map(row => String(row[0]).trim());
  const rowIndex = idList.indexOf(id);

  if (rowIndex > 0) {
    if (isNotes) {
      sheet.getRange(rowIndex + 1, 2, 1, 3).setValues([[e.parameter.kategori || "Umum", e.parameter.judul || "", e.parameter.isi || ""]]);
    } else {
      sheet.getRange(rowIndex + 1, 2, 1, 7).setValues([[e.parameter.kategori || "", e.parameter.nama_aplikasi || "", e.parameter.username || "", e.parameter.email || "", e.parameter.password || "", e.parameter.ket || "", e.parameter.ket_dua || ""]]);
    }
    
    if (e.parameter.favorit !== undefined) {
      const headers = data[0];
      const favCol = headers.indexOf("favorit") + 1;
      if (favCol > 0) {
        let favStatus = String(e.parameter.favorit).toLowerCase().trim() === 'true';
        sheet.getRange(rowIndex + 1, favCol).setValue(favStatus);
      }
    }
    
    if (e.parameter.vault_id) {
       const headers = data[0];
       const vaultIdCol = headers.indexOf("vault_id") + 1;
       if (vaultIdCol > 0) sheet.getRange(rowIndex + 1, vaultIdCol).setValue(e.parameter.vault_id);
    }
    
    SpreadsheetApp.flush();
    return output({ success: true, message: "updated" });
  }
  return output({ success: false, message: "ID not found" });
}

/* =====================================================
   🔄 UPDATE VAULT ID (Pindah Brankas)
===================================================== */
function updateVaultId(e, isNotes) {
  const targetTab = isNotes ? NOTES_SHEET_NAME : SHEET_NAME;
  const sheet = getSheet(targetTab);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const id = String(e.parameter.id || "").trim();
  const newVaultId = e.parameter.new_vault_id || "default";
  
  const vaultIdCol = headers.indexOf("vault_id") + 1;
  if (vaultIdCol === 1) return output({ success: false, message: "Kolom vault_id tidak ditemukan." });

  const idList = data.map(row => String(row[0]).trim());
  const rowIndex = idList.indexOf(id);

  if (rowIndex > 0) {
    sheet.getRange(rowIndex + 1, vaultIdCol).setValue(newVaultId);
    SpreadsheetApp.flush();
    return output({ success: true, message: "Vault ID diubah ke: " + newVaultId });
  }
  return output({ success: false, message: "ID not found" });
}

/* =====================================================
   📁 UPDATE KATEGORI
===================================================== */
function updateCategoryData(e, isNotes) {
  const targetTab = isNotes ? NOTES_SHEET_NAME : SHEET_NAME;
  const sheet = getSheet(targetTab);
  const data = sheet.getDataRange().getValues();
  const id = String(e.parameter.id || "").trim();
  const newKategori = e.parameter.new_kategori || "Umum";
  
  const idList = data.map(row => String(row[0]).trim());
  const rowIndex = idList.indexOf(id);

  if (rowIndex > 0) {
    sheet.getRange(rowIndex + 1, 2).setValue(newKategori);
    SpreadsheetApp.flush();
    return output({ success: true, message: "Kategori diubah ke: " + newKategori });
  }
  return output({ success: false, message: "ID not found" });
}

/* =====================================================
   🗑️ DELETE DATA
===================================================== */
function deleteData(e, isNotes) {
  const targetTab = isNotes ? NOTES_SHEET_NAME : SHEET_NAME;
  const sheet = getSheet(targetTab);
  const data = sheet.getDataRange().getValues();
  const id = String(e.parameter.id || "").trim();

  const idList = data.map(row => String(row[0]).trim());
  const rowIndex = idList.indexOf(id);

  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex + 1);
    SpreadsheetApp.flush();
    return output({ success: true, message: "deleted" });
  }
  return output({ success: false, message: "ID not found" });
}

/* =====================================================
   🗑️ BULK DELETE DATA (OPTIMIZED SPEED)
===================================================== */
function bulkDeleteData(e, isNotes){
   PropertiesService.getScriptProperties().setProperty("BULK_PROCESS_RUNNING", "true");
   
   try {
     const targetTab = isNotes ? NOTES_SHEET_NAME : SHEET_NAME;
     const sheet = getSheet(targetTab);
     const ids = JSON.parse(e.parameter.ids || "[]").map(id => String(id).trim());
     
     if(!Array.isArray(ids) || ids.length === 0){
        return output({success: false, message: "No IDs"});
     }
     
     const data = sheet.getDataRange().getValues();
     const headers = data[0];
     let totalTerhapus = 0;
     const dataTetap = [headers]; 
     
     for(let i = 1; i < data.length; i++){
        const currentId = String(data[i][0]).trim();
        if (ids.indexOf(currentId) > -1) {
          totalTerhapus++; 
        } else {
          dataTetap.push(data[i]); 
        }
     }
     
     if (totalTerhapus > 0) {
       sheet.clearContents(); 
       sheet.getRange(1, 1, dataTetap.length, dataTetap[0].length).setValues(dataTetap); 
       SpreadsheetApp.flush(); 
     }
     
     return output({success: true, deleted: totalTerhapus});
   } catch(err) {
     return output({ success: false, error: err.toString() });
   } finally {
     PropertiesService.getScriptProperties().setProperty("BULK_PROCESS_RUNNING", "false");
   }
}

/* =====================================================
   ⭐ TOGGLE FAVORITE (🚀 RESPONS HIGH-SPEED ACTION - FIXED)
===================================================== */
function toggleFavoriteData(e, isNotes) {
  const targetTab = isNotes ? NOTES_SHEET_NAME : SHEET_NAME;
  const sheet = getSheet(targetTab);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return output({ success: false, message: "Sheet kosong" });
  
  // Ambil data ID dengan pembersihan spasi menyeluruh
  const idList = sheet.getRange(1, 1, lastRow, 1).getValues().map(row => String(row[0]).trim());
  const id = String(e.parameter.id || "").trim();
  const rowIndex = idList.indexOf(id); 

  if (rowIndex > 0) { 
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const favColumnIndex = headers.indexOf("favorit") + 1;
    
    if (favColumnIndex === 0) return output({ success: false, message: "Kolom favorit tidak ada" });
    
    // Pembacaan dan penulisan sel tunggal yang presisi (rowIndex + 1)
    const targetCell = sheet.getRange(rowIndex + 1, favColumnIndex);
    const val = targetCell.getValue();
    
    const statusSekarang = (typeof val === "boolean") ? val : String(val || '').toLowerCase().trim() === 'true';
    const statusBaru = !statusSekarang;
    
    targetCell.setValue(statusBaru);
    SpreadsheetApp.flush(); 
    
    return output({ success: true, message: "updated", favorit: statusBaru });
  }
  return output({ success: false, message: "ID not found: " + id });
}

/* =====================================================
   ✅ PROCESS SYNC QUEUE
===================================================== */
function processSyncQueue(e) {
  try {
    const queueItems = JSON.parse(e.parameter.queue_items || "[]");
    if (!Array.isArray(queueItems) || queueItems.length === 0) {
      return output({ success: true, message: "Queue kosong", processed: 0 });
    }
    
    let results = { processed: 0, succeeded: 0, failed: 0, errors: [] };
    
    queueItems.forEach((item, idx) => {
      try {
        const { action, data_type, ...payload } = item;
        const isNotes = data_type === "notes";
        const fakeEvent = { parameter: { action, data_type, ...payload } };
        
        switch(action) {
          case "create": createData(fakeEvent, isNotes); break;
          case "update": updateData(fakeEvent, isNotes); break;
          case "delete": deleteData(fakeEvent, isNotes); break;
          case "toggle_favorite": toggleFavoriteData(fakeEvent, isNotes); break;
          case "update_category": updateCategoryData(fakeEvent, isNotes); break;
          case "update_vault": updateVaultId(fakeEvent, isNotes); break;
          default: throw new Error("Unknown action: " + action);
        }
        results.succeeded++;
      } catch(err) {
        results.failed++;
        results.errors.push({ index: idx, action: item.action, id: item.id, error: err.toString() });
      }
      results.processed++;
    });
    
    SpreadsheetApp.flush();
    return output({ success: results.failed === 0, message: `Processed ${results.processed} items`, data: results });
  } catch(err) {
    return output({ success: false, error: "Sync queue error: " + err.toString() });
  }
}

/* =====================================================
   📊 REORDER FOLDERS
===================================================== */
function reorderFolders(e) {
  try {
    const dataType = e.parameter.data_type || "vault";
    const folderOrder = JSON.parse(e.parameter.folder_order || "[]");
    const propKey = "folder_order_" + dataType;
    PropertiesService.getScriptProperties().setProperty(propKey, JSON.stringify(folderOrder));
    return output({ success: true, message: `Folder order saved for ${dataType}`, data: { dataType, count: folderOrder.length, order: folderOrder } });
  } catch(err) {
    return output({ success: false, error: "Reorder folders error: " + err.toString() });
  }
}

/* =====================================================
   🔧 ON EDIT TRIGGER (ANTI BREAK-LAG & INFINITE LOOP FIXED)
===================================================== */
function onEdit(e) {
  if (!e || !e.range) return; 

  const scriptProperties = PropertiesService.getScriptProperties();

  // BYPASS UTAMAKAN: Lewati jika proses bulk sedang jalan ATAU script sedang menulis modifikasi baris
  if (scriptProperties.getProperty("BULK_PROCESS_RUNNING") === "true" || 
      scriptProperties.getProperty("ON_EDIT_LOCK") === "true") {
    return;
  }

  try {
    const sheet = e.range.getSheet();
    const currentSheetName = sheet.getName();
    if (currentSheetName !== SHEET_NAME && currentSheetName !== NOTES_SHEET_NAME) return;

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    if (startRow === 1) return; 

    const isNotes = currentSheetName !== NOTES_SHEET_NAME;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const idCol = headers.indexOf("id");
    const favCol = headers.indexOf("favorit");
    const dateCol = headers.indexOf(isNotes ? "tanggal" : "created_at");
    const vaultIdCol = headers.indexOf("vault_id");
    
    const fullRange = sheet.getRange(startRow, 1, numRows, lastCol);
    const fullValues = fullRange.getValues();
    
    let updated = false;

    for (let r = 0; r < numRows; r++) {
      const hasContent = fullValues[r].some((v, idx) => idx !== idCol && v !== "");

      if (hasContent) {
        if (idCol > -1 && String(fullValues[r][idCol]).trim() === "") { fullValues[r][idCol] = generateID(isNotes ? "NOTE" : "VAULT"); updated = true; }
        if (dateCol > -1 && fullValues[r][dateCol] === "") { fullValues[r][dateCol] = new Date(); updated = true; }
        if (favCol > -1 && fullValues[r][favCol] === "") { fullValues[r][favCol] = false; updated = true; }
        if (vaultIdCol > -1 && fullValues[r][vaultIdCol] === "") { fullValues[r][vaultIdCol] = "default"; updated = true; }
      }
    }

    if (updated) {
      // Nyalakan pengunci lokal sebelum memanggil setValues untuk memutus rantai trigger berulang
      scriptProperties.setProperty("ON_EDIT_LOCK", "true");
      fullRange.setValues(fullValues);
      SpreadsheetApp.flush();
    }
  } catch(err) {
    Logger.log("❌ Error pada onEdit: " + err.toString());
  } finally {
    // Selalu pastikan pengunci dikembalikan ke false
    scriptProperties.setProperty("ON_EDIT_LOCK", "false");
  }
}
