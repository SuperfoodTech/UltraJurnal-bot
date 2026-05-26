/**
 * UltraJurnal.bot — Google Apps Script Upload Handler
 * =====================================================
 * Deploy sebagai Web App:
 *   - Execute as: Me (akun pemilik folder GDrive)
 *   - Who has access: Anyone
 *
 * Discord bot mengirim file via HTTP POST (multipart/form-data atau JSON base64)
 * ke URL Web App ini, lalu script menyimpan file ke folder GDrive dan
 * mengembalikan link shareable.
 *
 * Target folder: https://drive.google.com/drive/folders/17jmj2j6WjsyEgxBNV7XOK6GT_2FplaXy
 */

var FOLDER_ID = "17jmj2j6WjsyEgxBNV7XOK6GT_2FplaXy";

// ─────────────────────────────────────────────────────────────────────────────
// Respons helper
// ─────────────────────────────────────────────────────────────────────────────
function jsonResponse(data, statusCode) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ─────────────────────────────────────────────────────────────────────────────
// GET — Health check (untuk testing)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  return jsonResponse({
    status: "ok",
    message: "UltraJurnal.bot GDrive Upload Handler aktif ✅",
    folder_id: FOLDER_ID,
    timestamp: new Date().toISOString()
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// POST — Upload file ke Google Drive
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Payload JSON yang diterima dari Discord bot:
 * {
 *   "filename": "nama_file.xlsx",
 *   "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
 *   "data": "<base64 encoded file content>",
 *   "discord_user": "username#0000"   (opsional, untuk logging)
 * }
 *
 * Response:
 * {
 *   "success": true,
 *   "link": "https://drive.google.com/file/d/FILE_ID/view?usp=sharing",
 *   "file_id": "FILE_ID",
 *   "filename": "nama_file.xlsx"
 * }
 */
function doPost(e) {
  try {
    // ── Parse payload ─────────────────────────────────────────────────────
    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse({ success: false, error: "Payload JSON tidak valid: " + parseErr.message });
    }

    var filename    = payload.filename    || ("upload_" + new Date().getTime() + ".xlsx");
    var mimeType    = payload.mimeType    || "application/octet-stream";
    var base64Data  = payload.data;
    var discordUser = payload.discord_user || "unknown";

    if (!base64Data) {
      return jsonResponse({ success: false, error: "Field 'data' (base64) tidak ada dalam payload." });
    }

    // ── Decode base64 → Blob ──────────────────────────────────────────────
    var decoded  = Utilities.base64Decode(base64Data);
    var blob     = Utilities.newBlob(decoded, mimeType, filename);

    // ── Dapatkan folder tujuan ────────────────────────────────────────────
    var folder;
    try {
      folder = DriveApp.getFolderById(FOLDER_ID);
    } catch (folderErr) {
      return jsonResponse({ success: false, error: "Folder GDrive tidak ditemukan: " + FOLDER_ID });
    }

    // ── Upload file ke folder ─────────────────────────────────────────────
    var file = folder.createFile(blob);

    // ── Set permission: anyone can view ───────────────────────────────────
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId   = file.getId();
    var fileLink = "https://drive.google.com/file/d/" + fileId + "/view?usp=sharing";

    Logger.log("✅ Upload berhasil: " + filename + " oleh " + discordUser + " → " + fileLink);

    return jsonResponse({
      success:  true,
      link:     fileLink,
      file_id:  fileId,
      filename: filename
    });

  } catch (err) {
    Logger.log("❌ Error: " + err.message + "\n" + err.stack);
    return jsonResponse({
      success: false,
      error:   err.message
    });
  }
}
