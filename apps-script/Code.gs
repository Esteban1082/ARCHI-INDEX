const SHEET_NAME = 'Garanties';
const DRIVE_FOLDER_NAME = 'Garantie Keeper - tickets';
const HEADERS = ['ID','Nom','Catégorie','Date achat','Durée garantie mois','Date fin garantie','Statut','Magasin','Commande','Prix','Ticket','Notes','Mise à jour'];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    if (action === 'list') return json({ ok: true, items: listItems() });
    if (action === 'save') return json(saveItem(body.item, body.file));
    if (action === 'delete') return json(deleteItem(body.id));
    return json({ ok: false, error: 'Action inconnue' });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}

function doGet() {
  return json({ ok: true, items: listItems() });
}

function listItems() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(row => row[0]).map(row => ({
    id: row[0],
    name: row[1],
    category: row[2],
    purchaseDate: toIsoDate(row[3]),
    warrantyMonths: Number(row[4] || 0),
    seller: row[7],
    orderNumber: row[8],
    price: Number(row[9] || 0),
    receiptUrl: row[10],
    notes: row[11]
  }));
}

function saveItem(item, file) {
  const sheet = getSheet();
  const receiptUrl = file ? uploadReceipt(file, item.id) : item.receiptUrl;
  const row = [
    item.id,
    item.name,
    item.category,
    item.purchaseDate,
    Number(item.warrantyMonths || 0),
    warrantyEnd(item.purchaseDate, item.warrantyMonths),
    warrantyStatus(item.purchaseDate, item.warrantyMonths),
    item.seller,
    item.orderNumber,
    Number(item.price || 0),
    receiptUrl || '',
    item.notes,
    new Date()
  ];
  const index = findRowById(sheet, item.id);
  if (index > 0) sheet.getRange(index, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return { ok: true, receiptUrl: receiptUrl || '' };
}

function deleteItem(id) {
  const sheet = getSheet();
  const index = findRowById(sheet, id);
  if (index > 0) sheet.deleteRow(index);
  return { ok: true };
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function findRowById(sheet, id) {
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < values.length; i++) if (values[i][0] === id) return i + 1;
  return -1;
}

function uploadReceipt(file, id) {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  const bytes = Utilities.base64Decode(file.data);
  const safeName = `${id}-${file.name}`.replace(/[\\/:*?"<>|]/g, '-');
  const blob = Utilities.newBlob(bytes, file.type || 'application/octet-stream', safeName);
  const driveFile = folder.createFile(blob);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return driveFile.getUrl();
}

function warrantyEnd(purchaseDate, months) {
  const date = new Date(purchaseDate);
  date.setMonth(date.getMonth() + Number(months || 0));
  return toIsoDate(date);
}

function warrantyStatus(purchaseDate, months) {
  const days = Math.ceil((new Date(warrantyEnd(purchaseDate, months)) - new Date()) / 86400000);
  if (days < 0) return 'Expiré';
  if (days <= 30) return 'Bientôt expiré';
  return 'OK';
}

function toIsoDate(value) {
  if (!value) return '';
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function json(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
