// ============================================
// Job Card Details - Entry Form (Web App)
// Google Apps Script - Supersonic Group
// ============================================
//
// FUTURE ME CHANGES KARNE KE LIYE (isko hi edit karo, HTML nahi):
//   - Naya field add/remove/reorder karna ho  -> neeche FIELDS array edit karo
//   - Sheet ka tab naam badalna ho            -> CONFIG.SHEET_NAME edit karo
//   - Unit ki dropdown suggestions badalni ho -> UNIT_SUGGESTIONS array edit karo
//
// FIELDS array hi form (kaunse box dikhenge) aur Sheet me column (kis order
// me data save hoga) dono control karta hai. Ek field add karoge to woh
// form me apne aap dikhega - bas Sheet me bhi wahi naam ka column add karna
// mat bhoolna (FIELDS me jis order me hai, Sheet me A,B,C... us order me hi
// hone chahiye, S.No column ke baad).

const CONFIG = {
  SHEET_NAME: "Job Card Details",   // Sheet tab ka naam. Match na mile to script apne aap pehli tab use kar legi.
  FORM_TITLE: "Job Card Details - Entry Form",
  COMPANY_NAME: "Supersonic Group"
};

// Har field: key = Sheet column ka header + data ka naam, label = form pe dikhega,
// type = "text" | "textarea" | "number" | "date", required = true/false,
// voice = true matlab mic button milega (bolke type kar sakte ho).
const FIELDS = [
  { key: "date",         label: "Date",                         type: "date",     required: true,  voice: false },
  { key: "jobCardNo",    label: "Job Card No.",                  type: "text",     required: false,  voice: true  },
  { key: "machineName",  label: "Machine Name",                  type: "text",     required: false,  voice: true  },
  { key: "machineSerial",label: "Machine Serial No.",            type: "text",     required: false, voice: true  },
  { key: "materialName", label: "Material / Spare Part Name",    type: "text",     required: false,  voice: true  },
  { key: "materialCode", label: "Material Code",                 type: "text",     required: false, voice: true  },
  { key: "qtyUsed",      label: "Qty Used",                      type: "number",   required: false,  voice: true  },
  { key: "unit",         label: "Unit",                          type: "text",     required: false, voice: true  },
  { key: "technician",   label: "Technician",                    type: "text",     required: false,  voice: true  },
  { key: "remarks",      label: "Remarks",                       type: "textarea", required: false, voice: true  }
];

// Unit box me typing karte waqt ye suggestions dikhengi (dropdown ki tarah,
// par phir bhi khud kuch bhi type/bol sakte ho).
const UNIT_SUGGESTIONS = ["Pcs", "Nos", "Set", "Meter", "Kg", "Litre", "Box", "Roll"];

// ============================================
// WEB APP ENTRY POINTS
// ============================================
//
// Ye script 2 tarah se use hoti hai:
//  1) Seedha /exec URL kholo -> yahi doGet() wala form dikhta hai (typing works,
//     par is form ke andar voice/mic Google ki apni sandboxing ki wajah se kaam
//     nahi karta - Google ki limitation hai, hamare code ki nahi).
//  2) standalone/job-card-form.html (GitHub Pages pe host hoti hai) -> wahan
//     mic bhi kaam karta hai, kyunki wo Google ke sandbox ke bahar hai. Wo page
//     is hi script se doGet(?action=fields) se form-fields mangata hai aur
//     doPost() se data save karta hai.

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "fields") {
    return jsonOutput_({ fields: FIELDS, units: UNIT_SUGGESTIONS, formTitle: CONFIG.FORM_TITLE, companyName: CONFIG.COMPANY_NAME });
  }

  var template = HtmlService.createTemplateFromFile("Form");
  template.formTitle = CONFIG.FORM_TITLE;
  template.companyName = CONFIG.COMPANY_NAME;
  template.fieldsJson = JSON.stringify(FIELDS);
  template.unitSuggestionsJson = JSON.stringify(UNIT_SUGGESTIONS);

  return template.evaluate()
    .setTitle(CONFIG.FORM_TITLE)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// GitHub Pages wala standalone form isi ko call karta hai (fetch POST) data save karne ke liye.
function doPost(e) {
  try {
    var formData = JSON.parse(e.postData.contents);
    var result = saveEntry_(formData);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// HTML templating ke liye - Style.html / ClientScript.html ko Form.html me include karta hai.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// SHEET HELPER
// ============================================
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
}

// ============================================
// FORM SUBMIT
// ============================================
// google.script.run se (Apps Script ke apne /exec form se) yahi call hota hai.
function submitForm(formData) {
  return saveEntry_(formData);
}

// Dono submitForm() aur doPost() (GitHub Pages wala standalone form) yahi
// shared logic use karte hai, taaki validation/Sheet-save ek hi jagah ho.
function saveEntry_(formData) {
  var sheet = getSheet_();

  // Required fields check.
  for (var i = 0; i < FIELDS.length; i++) {
    var f = FIELDS[i];
    if (f.required && (!formData[f.key] || String(formData[f.key]).trim() === "")) {
      throw new Error(f.label + " zaroori hai (required).");
    }
  }

  var lastRow = sheet.getLastRow();
  var nextSerial = lastRow < 1 ? 1 : lastRow; // header row 1 hai, to data rows = lastRow - 1, next serial = lastRow

  // Row banate hain: S.No pehla column, phir FIELDS ke order me baaki columns.
  var row = [nextSerial];
  FIELDS.forEach(function (f) {
    var value = formData[f.key] || "";
    if (f.type === "date" && value) {
      // yyyy-mm-dd string aata hai date input se, Sheet me readable date store karo.
      var parts = value.split("-");
      value = new Date(parts[0], parts[1] - 1, parts[2]);
    }
    row.push(value);
  });

  sheet.appendRow(row);

  return {
    success: true,
    serialNo: nextSerial
  };
}
