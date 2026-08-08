// Ganti dengan URL deployment Apps Script Anda
const SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbzuoCCKWlyWp4OXrBvgIYi2J3t5HxVoWrCtgfHZ4DCW525aFFOnMewmzyE4fsLM5FtM/exec";
const MARK_HEADER_START = "PARTS LIST INQUIRY";
const MARK_ITEM_HEADER = "ITEM NUMBER";
const MARK_BODY_END = "MAIN MENU";
const LABEL_TOY_NUMBER = "TOY NUMBER";
const LABEL_LAST_CHANGE = "LAST CHANGE DATE";
const LABEL_APPROVAL = "APPROVAL DATE";

// --- Helper: setara ExtractToyNumber di VBA ---
function extractToyNumber(raw) {
  const upper = raw.toUpperCase();
  const p = upper.indexOf(LABEL_TOY_NUMBER);
  if (p === -1) return "";
  const posColon = raw.indexOf(":", p);
  if (posColon === -1) return "";
  const tail = raw.substring(posColon + 1);
  const tailUpper = tail.toUpperCase();
  const posParen = tail.indexOf("(");
  const posCreated = tailUpper.indexOf("CREATED DATE");

  let cutPos;
  if (posParen !== -1 && (posCreated === -1 || posParen < posCreated)) {
    cutPos = posParen;
  } else if (posCreated !== -1) {
    cutPos = posCreated;
  } else {
    cutPos = tail.length;
  }

  let hasil = tail.substring(0, cutPos).trim();
  hasil = hasil.replace(/\s{2,}/g, " ");
  return hasil;
}

// --- Helper: setara ExtractDateAfterLabel di VBA ---
function extractDateAfterLabel(raw, label) {
  const upper = raw.toUpperCase();
  const p = upper.indexOf(label);
  if (p === -1) return "";
  const posColon = raw.indexOf(":", p);
  if (posColon === -1) return "";
  const tail = raw.substring(posColon + 1);
  const tailUpper = tail.toUpperCase();

  let posBy = tailUpper.indexOf("BY:");
  if (posBy === -1) posBy = tailUpper.indexOf("BY ");

  const cutPos = posBy !== -1 ? posBy : tail.length;
  return tail.substring(0, cutPos).trim();
}

// --- Helper: setara TryParsePartRow di VBA ---
function tryParsePartRow(line) {
  let cleaned = line.trim().replace(/\s{2,}/g, " ");
  if (cleaned.length === 0) return null;

  const parts = cleaned.split(" ");
  if (parts.length < 2) return null;

  const tok0 = parts[0];
  if (tok0.indexOf("-") === -1) return null;
  if (!/[A-Za-z0-9]/.test(tok0)) return null;

  const pn = tok0;
  let idx = 1;
  if (idx < parts.length && parts[idx] === "*") idx++;
  const yld = idx < parts.length ? parts[idx] : "";

  return { pn, yld };
}

// --- Parser utama: port 1:1 dari state machine ExtractPartList di VBA ---
function parseRawText(rawText) {
  const lines = rawText.split(/\r?\n/);
  const outputRows = []; // item: null (baris kosong pemisah) atau {dashcode, note, partNumber, description, yield}

  let state = 0;
  let dashcode = "",
    strLastChange = "",
    strApproval = "";
  let perBlockRowCount = 0;
  let lastRowIndex = -1;

  for (const rawLineOriginal of lines) {
    const raw = rawLineOriginal.replace(/\u00A0/g, " "); // ganti non-breaking space
    const trimmed = raw.trim();
    const upper = raw.toUpperCase();

    switch (state) {
      case 0:
        if (upper.includes(MARK_HEADER_START)) {
          state = 1;
          dashcode = "";
          strLastChange = "";
          strApproval = "";
          perBlockRowCount = 0;
        }
        break;

      case 1:
        if (upper.includes(LABEL_TOY_NUMBER)) {
          dashcode = extractToyNumber(raw);
        } else if (upper.includes(LABEL_LAST_CHANGE)) {
          strLastChange =
            "LAST CHANGE DATE : " +
            extractDateAfterLabel(raw, LABEL_LAST_CHANGE);
        } else if (upper.includes(LABEL_APPROVAL)) {
          strApproval =
            "APPROVAL DATE : " +
            extractDateAfterLabel(raw, LABEL_APPROVAL);
        } else if (upper.includes(MARK_ITEM_HEADER)) {
          state = 2;
        }
        break;

      case 2:
        if (upper.includes(MARK_BODY_END)) {
          if (
            perBlockRowCount === 1 &&
            strApproval &&
            lastRowIndex !== -1
          ) {
            outputRows[lastRowIndex].note =
              (outputRows[lastRowIndex].note || "") + "\n" + strApproval;
          }
          state = 3;
        } else {
          const parsed = tryParsePartRow(trimmed);
          if (parsed) {
            const row = {
              dashcode: "",
              note: "",
              partNumber: parsed.pn,
              description: "",
              yield: parsed.yld,
            };
            if (perBlockRowCount === 0) {
              row.dashcode = dashcode.trim().replace(/\s/g, "");
              row.note = strLastChange;
            } else if (perBlockRowCount === 1) {
              row.note = strApproval;
            }
            outputRows.push(row);
            lastRowIndex = outputRows.length - 1;
            perBlockRowCount++;
          }
        }
        break;

      case 3:
        if (upper.includes(MARK_HEADER_START)) {
          outputRows.push(null); // baris kosong pemisah antar blok
          state = 1;
          dashcode = "";
          strLastChange = "";
          strApproval = "";
          perBlockRowCount = 0;
        }
        break;
    }
  }

  return outputRows;
}

function processAndSend() {
  const btn = document.getElementById("sendBtn");
  const status = document.getElementById("status");
  const rawText = document.getElementById("rawInput").value;

  const outputRows = parseRawText(rawText);
  const dataRowCount = outputRows.filter((r) => r !== null).length;

  if (dataRowCount === 0) {
    alert("Tidak ada data barang yang terdeteksi!");
    return;
  }

  // Susun jadi array 2D: [dashcode, note, partNumber, description, yield]
  const values = outputRows.map((r) =>
    r === null
      ? ["", "", "", "", ""]
      : [r.dashcode, r.note, r.partNumber, r.description, r.yield],
  );

  btn.disabled = true;
  status.textContent = "Mengirim " + dataRowCount + " baris data...";

  fetch(SCRIPT_URL, {
    method: "POST",
    mode: "no-cors",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(values),
  })
    .then(() => {
      status.textContent =
        "Berhasil mengirim " +
        dataRowCount +
        " baris data ke Google Sheets!";
      document.getElementById("rawInput").value = "";
    })
    .catch((err) => {
      status.textContent = "Gagal: " + err;
    })
    .finally(() => {
      btn.disabled = false;
    });
}
