const CHALLENGES = [
  { id: "selfie-noivos", title: "Selfie com os noivos", description: "Registre uma foto ao lado da Ana Clara e do Felipe — peça autorização antes." },
  { id: "pista-danca", title: "Passo de dança", description: "Mostre seu passo mais animado na pista de dança." },
  { id: "mesa-completa", title: "Mesa completa", description: "Reúna as pessoas da sua mesa para uma foto." },
  { id: "brinde-amigos", title: "Brinde entre amigos", description: "Registre um brinde especial com seus amigos." },
  { id: "tres-geracoes", title: "Três gerações", description: "Junte três gerações da família em uma foto." },
  { id: "coracao-maos", title: "Coração com as mãos", description: "Forme um coração com as mãos junto de alguém da festa." },
  { id: "nova-amizade", title: "Nova amizade", description: "Tire uma foto com alguém que você conheceu hoje — com consentimento." },
  { id: "look-festa", title: "Look da festa", description: "Mostre um detalhe especial do seu look ou de um amigo, com autorização." },
  { id: "danca-em-dupla", title: "Dança em dupla", description: "Convide alguém para dançar e registrem esse momento." },
  { id: "sorriso-espontaneo", title: "Sorriso espontâneo", description: "Capture um momento alegre da festa, sempre respeitando quem aparece na foto." },
  { id: "foto-com-familia", title: "Foto com a família", description: "Reúna familiares para uma foto alegre." },
  { id: "pose-criativa", title: "Pose criativa", description: "Crie uma pose divertida com alguém da festa." },
  { id: "brinde-especial", title: "Brinde especial", description: "Faça um brinde aos noivos e registre o momento." },
  { id: "detalhe-decoracao", title: "Detalhe da decoração", description: "Fotografe um detalhe bonito da decoração." },
  { id: "foto-com-amigo", title: "Foto com um amigo", description: "Registre a festa ao lado de um amigo querido." },
  { id: "convite-danca", title: "Convite para dançar", description: "Convide alguém para dançar e tirem uma foto juntos." },
  { id: "grupo-animado", title: "Grupo animado", description: "Monte um pequeno grupo e façam uma pose de comemoração." },
  { id: "detalhe-bolo", title: "Detalhe da celebração", description: "Registre um detalhe especial da mesa ou da festa." },
  { id: "passo-sincronizado", title: "Passo sincronizado", description: "Combine um passo de dança com alguém e registrem." },
  { id: "foto-alegria", title: "Alegria da festa", description: "Registre um momento feliz, com autorização de quem aparece." }
];
const CHALLENGE_EVENT_DATE = "2026-11-16";
const MAX_CHALLENGE_THUMBNAIL_BYTES = 120000;
const WEDDING_GALLERY_FOLDER_PROPERTY = "WEDDING_PUBLIC_GALLERY_FOLDER_ID";
const WEDDING_GALLERY_PAGE_SIZE = 48;

function isChallengeEventDay() {
  return Utilities.formatDate(new Date(), "America/Sao_Paulo", "yyyy-MM-dd") === CHALLENGE_EVENT_DATE;
}

function challengeLockedResponse() {
  return { locked: true, message: "Os desafios serão liberados somente no dia da festa." };
}

function assertChallengeEventDay() {
  if (!isChallengeEventDay()) throw new Error(challengeLockedResponse().message);
}

function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Serviço do casamento ativo.");
  }
  const action = e.parameter.action;
  const callback = e.parameter.callback || "";
  if (!/^[A-Za-z_$][A-Za-z0-9_$]{0,80}$/.test(callback)) {
    throw new Error("Callback inválido.");
  }

  let result;
  try {
    if (action === "weddingGallery") {
      result = getWeddingGalleryPage(e.parameter.pageToken || "");
    } else if (!isChallengeEventDay()) {
      result = challengeLockedResponse();
    } else if (action === "guestChallenges") {
      result = getGuestChallenges(e.parameter.name);
    } else if (action === "challengeProgress") {
      result = getCollectiveChallengeProgress();
    } else if (action === "guestCollage") {
      result = getGuestCollage(e.parameter.name);
    } else {
      throw new Error("Ação inválida.");
    }
  } catch (error) {
    result = { error: error instanceof Error ? error.message : "Não foi possível carregar os desafios." };
  }
  const json = JSON.stringify(result).replace(/</g, "\\u003c");
  return ContentService.createTextOutput(`${callback}(${json});`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getWeddingGalleryPage(pageToken) {
  if (typeof pageToken !== "string" || pageToken.length > 10000) {
    throw new Error("Token de página inválido.");
  }
  const folderId = PropertiesService.getScriptProperties()
    .getProperty(WEDDING_GALLERY_FOLDER_PROPERTY);
  if (!folderId) return { configured: false, photos: [], nextPageToken: "" };

  const files = pageToken
    ? DriveApp.continueFileIterator(pageToken)
    : DriveApp.getFolderById(folderId).getFiles();
  const photos = [];
  while (files.hasNext() && photos.length < WEDDING_GALLERY_PAGE_SIZE) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if (!mimeType.startsWith("image/") || mimeType === "image/svg+xml" || mimeType === "image/gif") {
      continue;
    }
    photos.push({
      id: file.getId(),
      name: file.getName().slice(0, 200),
      createdAt: Utilities.formatDate(file.getDateCreated(), "America/Sao_Paulo", "dd/MM/yyyy")
    });
  }
  return {
    configured: true,
    photos,
    nextPageToken: files.hasNext() ? files.getContinuationToken() : ""
  };
}

function setupWeddingGallery() {
  const properties = PropertiesService.getScriptProperties();
  let folderId = properties.getProperty(WEDDING_GALLERY_FOLDER_PROPERTY);
  let folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    folder = DriveApp.createFolder("Memorial do casamento - Ana Clara e Felipe");
    folderId = folder.getId();
    properties.setProperty(WEDDING_GALLERY_FOLDER_PROPERTY, folderId);
  }
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log(`Pasta pública do memorial: ${folder.getUrl()}`);
  return folder.getUrl();
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === "uploadIndividualChallengePhoto") return uploadIndividualChallengePhoto(data);

  const name = typeof data.name === "string" ? data.name.trim() : "";
  const attendance = data.attendance;
  const people = Number(data.people);
  const guests = Array.isArray(data.guests)
    ? data.guests.filter((guest) => typeof guest === "string").map((guest) => guest.trim()).filter(Boolean)
    : [];

  if (!name || name.length > 120) throw new Error("Nome inválido.");
  if (attendance !== "Sim, estarei lá" && attendance !== "Infelizmente não poderei ir") {
    throw new Error("Resposta de presença inválida.");
  }
  if (!Number.isInteger(people) || people < 1 || people > 10) {
    throw new Error("Quantidade de pessoas inválida.");
  }
  if (guests.some((guest) => guest.length > 120)) throw new Error("Nome de acompanhante muito longo.");
  if (attendance === "Sim, estarei lá" && guests.length !== people - 1) {
    throw new Error("A quantidade de acompanhantes não corresponde ao total de pessoas.");
  }
  if (attendance === "Sim, estarei lá") {
    const normalizedNames = [name, ...guests].map(normalizeGuestName);
    if (new Set(normalizedNames).size !== normalizedNames.length) {
      throw new Error("Cada pessoa precisa de um nome completo diferente.");
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const sheet = getResponsesSheet(spreadsheet);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const guestColumn = ensureGuestColumn(sheet);
    const row = [new Date(), safeCellText(name), safeCellText(attendance), people];
    while (row.length < guestColumn - 1) row.push("");
    row.push(safeCellText(guests.join("\n")));
    sheet.appendRow(row);
    if (attendance === "Sim, estarei lá") {
      const registry = getRegisteredGuestsSheet(spreadsheet);
      [name, ...guests].forEach(guestName => ensureRegisteredGuest(registry, guestName));
    }
  } finally {
    lock.releaseLock();
  }

  return ContentService.createTextOutput("OK");
}

function setupGuestColumn() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const sheet = getResponsesSheet(spreadsheet);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const column = ensureGuestColumn(sheet);
    Logger.log(`Coluna Acompanhantes criada/verificada na aba "${sheet.getName()}", coluna ${column}.`);
  } finally {
    lock.releaseLock();
  }
}

function getResponsesSheet(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    if (sheet.getLastRow() === 0) continue;
    const columnCount = Math.max(sheet.getLastColumn(), 4);
    const headers = sheet.getRange(1, 1, 1, columnCount).getDisplayValues()[0]
      .map((header) => header.trim());
    if (headers[0] === "Data e hora" && headers[1] === "Nome" &&
        headers[2] === "Presença" && headers[3] === "Pessoas") {
      return sheet;
    }
  }
  return spreadsheet.getActiveSheet();
}

function ensureGuestColumn(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Data e hora", "Nome", "Presença", "Pessoas", "Acompanhantes"]);
    return 5;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 4);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const existingGuestColumn = headers.indexOf("Acompanhantes");
  if (existingGuestColumn >= 0) return existingGuestColumn + 1;

  const oldNotesColumn = headers.indexOf("Restrição alimentar ou recado");
  const blankColumn = headers.findIndex((header, index) => index >= 4 && !header);
  const guestColumn = oldNotesColumn >= 0 || blankColumn < 0
    ? lastColumn + 1
    : blankColumn + 1;
  sheet.getRange(1, guestColumn).setValue("Acompanhantes");
  return guestColumn;
}

function safeCellText(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function uploadIndividualChallengePhoto(data) {
  assertChallengeEventDay();
  const guestName = typeof data.guestName === "string" ? data.guestName.trim() : "";
  const normalizedName = normalizeGuestName(guestName);
  if (!normalizedName || guestName.length > 120) throw new Error("Nome de convidado inválido.");
  if (!CHALLENGES.some(challenge => challenge.id === data.challengeId)) throw new Error("Desafio inválido.");
  const mimeType = data.mimeType;
  const extensions = { "image/jpeg": "jpg", "image/png": "png", "image/webp": "webp" };
  if (!Object.prototype.hasOwnProperty.call(extensions, mimeType)) throw new Error("Formato de foto inválido.");
  if (typeof data.base64 !== "string" || data.base64.length > 1400000 ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(data.base64)) {
    throw new Error("Arquivo de foto inválido ou muito grande.");
  }
  const bytes = Utilities.base64Decode(data.base64);
  if (bytes.length > 1000000) throw new Error("A foto precisa ter menos de 1 MB após a compressão.");
  if (!hasValidImageSignature(bytes, mimeType)) throw new Error("O conteúdo enviado não corresponde a uma imagem válida.");
  let thumbnailBytes = null;
  if (data.thumbnailBase64 !== undefined) {
    if (typeof data.thumbnailBase64 !== "string" || data.thumbnailBase64.length > 160000 ||
        !/^[A-Za-z0-9+/]+={0,2}$/.test(data.thumbnailBase64)) {
      throw new Error("Miniatura inválida ou muito grande.");
    }
    thumbnailBytes = Utilities.base64Decode(data.thumbnailBase64);
    if (thumbnailBytes.length > MAX_CHALLENGE_THUMBNAIL_BYTES ||
        !hasValidImageSignature(thumbnailBytes, "image/jpeg")) {
      throw new Error("O conteúdo da miniatura não corresponde a uma imagem JPEG válida.");
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const guest = findRegisteredGuest(getRegisteredGuestsSheet(spreadsheet), normalizedName);
    if (!guest) throw new Error("Nome não encontrado. Confirme presença antes de entrar nos desafios.");
    if (!guest.challengeIds.includes(data.challengeId)) throw new Error("Este desafio não foi atribuído a esse convidado.");
    const sheet = getIndividualSubmissionsSheet(spreadsheet);
    const existing = getGuestCompletedChallengeIds(sheet, normalizedName);
    if (existing.includes(data.challengeId)) throw new Error("Esse desafio já foi concluído.");
    const fileName = `${normalizedName.slice(0, 40)}-${data.challengeId}-${Date.now()}.${extensions[mimeType]}`;
    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const folder = getPrivateChallengePhotosFolder();
    const file = folder.createFile(blob);
    let thumbnailFile = null;
    try {
      if (thumbnailBytes) {
        const thumbnailName = `${normalizedName.slice(0, 40)}-${data.challengeId}-${Date.now()}-thumb.jpg`;
        thumbnailFile = folder.createFile(
          Utilities.newBlob(thumbnailBytes, "image/jpeg", thumbnailName)
        );
      }
      sheet.appendRow([
        new Date(),
        safeCellText(normalizedName),
        data.challengeId,
        file.getId(),
        thumbnailFile ? thumbnailFile.getId() : ""
      ]);
    } catch (error) {
      file.setTrashed(true);
      if (thumbnailFile) thumbnailFile.setTrashed(true);
      throw error;
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("OK");
}

function hasValidImageSignature(bytes, mimeType) {
  const first = (index) => bytes[index] & 0xff;
  if (mimeType === "image/jpeg") return first(0) === 0xff && first(1) === 0xd8 && first(2) === 0xff;
  if (mimeType === "image/png") {
    return [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
      .every((value, index) => first(index) === value);
  }
  if (mimeType === "image/webp") {
    return String.fromCharCode(...bytes.slice(0, 4)) === "RIFF" &&
      String.fromCharCode(...bytes.slice(8, 12)) === "WEBP";
  }
  return false;
}

function getGuestChallenges(name) {
  assertChallengeEventDay();
  const normalizedName = normalizeGuestName(name);
  if (!normalizedName || normalizedName.length > 120) throw new Error("Digite o nome completo usado na confirmação.");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const guest = findRegisteredGuest(getRegisteredGuestsSheet(spreadsheet), normalizedName);
  if (!guest) throw new Error("Nome não encontrado. Confirme presença antes de entrar nos desafios.");
  const completedChallengeIds = getGuestCompletedChallengeIds(
    getIndividualSubmissionsSheet(spreadsheet),
    normalizedName
  );
  const overallProgress = getCollectiveChallengeProgress();
  return {
    name: guest.name,
    challenges: guest.challengeIds.map(id => CHALLENGES.find(challenge => challenge.id === id)),
    completedChallengeIds,
    overallProgress
  };
}

function getGuestCollage(name) {
  assertChallengeEventDay();
  const normalizedName = normalizeGuestName(name);
  if (!normalizedName || normalizedName.length > 120) throw new Error("Digite o nome completo usado na confirmação.");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const guest = findRegisteredGuest(getRegisteredGuestsSheet(spreadsheet), normalizedName);
  if (!guest) throw new Error("Nome não encontrado. Confirme presença antes de ver a colagem.");

  const sheet = getIndividualSubmissionsSheet(spreadsheet);
  const lastRow = sheet.getLastRow();
  const submissions = lastRow < 2
    ? []
    : sheet.getRange(2, 2, lastRow - 1, 4).getValues()
      .filter(row => normalizeGuestName(row[0]) === normalizedName);
  const byChallenge = new Map(submissions.map(row => [row[1], { originalId: row[2], thumbnailId: row[3] }]));
  const completedChallengeIds = guest.challengeIds.filter(id => byChallenge.has(id));
  if (completedChallengeIds.length !== guest.challengeIds.length) {
    return { ready: false, completedCount: completedChallengeIds.length, requiredCount: guest.challengeIds.length };
  }

  const photos = guest.challengeIds.map(challengeId => {
    const submission = byChallenge.get(challengeId);
    const fileId = submission.thumbnailId || submission.originalId;
    if (!fileId) throw new Error("Uma das fotos do convidado não foi encontrada no Drive.");
    const blob = DriveApp.getFileById(fileId).getBlob();
    const mimeType = blob.getContentType();
    const bytes = blob.getBytes();
    if (!["image/jpeg", "image/png", "image/webp"].includes(mimeType) || bytes.length > 1000000) {
      throw new Error("Uma das fotos do convidado não pode ser incluída na colagem.");
    }
    return {
      challengeId,
      mimeType,
      base64: Utilities.base64Encode(bytes)
    };
  });
  return { ready: true, photos };
}

function getCollectiveChallengeProgress() {
  assertChallengeEventDay();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const guestsSheet = getRegisteredGuestsSheet(spreadsheet);
  const guests = Math.max(0, guestsSheet.getLastRow() - 1);
  const total = guests * 10;
  const submissions = getIndividualSubmissionsSheet(spreadsheet);
  const lastRow = submissions.getLastRow();
  const completed = lastRow < 2
    ? 0
    : new Set(submissions.getRange(2, 2, lastRow - 1, 2).getValues()
      .map(row => `${normalizeGuestName(row[0])}|${row[1]}`)).size;
  return { percent: total ? Math.min(100, Math.round(completed / total * 100)) : 0 };
}

function getRegisteredGuestsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName("Convidados dos desafios");
  if (!sheet) sheet = spreadsheet.insertSheet("Convidados dos desafios");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Nome", "Nome normalizado", "Desafios sorteados", "Data de cadastro"]);
  }
  return sheet;
}

function getIndividualSubmissionsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName("Envios dos desafios");
  if (!sheet) sheet = spreadsheet.insertSheet("Envios dos desafios");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Data e hora", "Nome normalizado", "Desafio", "ID privado no Drive", "ID privado da miniatura no Drive"]);
  } else {
    const thumbnailHeader = sheet.getLastColumn() >= 5
      ? String(sheet.getRange(1, 5).getDisplayValue()).trim()
      : "";
    if (!thumbnailHeader) {
      sheet.getRange(1, 5).setValue("ID privado da miniatura no Drive");
    } else if (thumbnailHeader !== "ID privado da miniatura no Drive") {
      throw new Error("A coluna 5 da aba Envios dos desafios precisa ser 'ID privado da miniatura no Drive'.");
    }
  }
  return sheet;
}

function ensureRegisteredGuest(sheet, name) {
  const normalizedName = normalizeGuestName(name);
  if (findRegisteredGuest(sheet, normalizedName)) return;
  const challengeIds = CHALLENGES.map(challenge => challenge.id);
  for (let index = challengeIds.length - 1; index > 0; index--) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [challengeIds[index], challengeIds[swapIndex]] = [challengeIds[swapIndex], challengeIds[index]];
  }
  sheet.appendRow([
    safeCellText(name),
    safeCellText(normalizedName),
    JSON.stringify(challengeIds.slice(0, 10)),
    new Date()
  ]);
}

function findRegisteredGuest(sheet, normalizedName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
    .filter(row => normalizeGuestName(row[1]) === normalizedName);
  if (rows.length > 1) throw new Error("Há nomes repetidos. Peça aos noivos para diferenciar seu nome completo.");
  if (!rows.length) return null;
  const challengeIds = JSON.parse(rows[0][2]);
  if (!Array.isArray(challengeIds) || challengeIds.length !== 10 ||
      challengeIds.some(id => !CHALLENGES.some(challenge => challenge.id === id))) {
    throw new Error("A lista de desafios deste convidado está inválida.");
  }
  return { name: rows[0][0], challengeIds };
}

function getGuestCompletedChallengeIds(sheet, normalizedName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 2, lastRow - 1, 2).getDisplayValues();
  return [...new Set(rows
    .filter(row => normalizeGuestName(row[0]) === normalizedName)
    .map(row => row[1])
    .filter(id => CHALLENGES.some(challenge => challenge.id === id)))];
}

function normalizeGuestName(name) {
  return String(name || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().replace(/\s+/g, " ").trim();
}

function getPrivateChallengePhotosFolder() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty("WEDDING_PRIVATE_CHALLENGE_PHOTOS_FOLDER_ID");
  if (folderId) return DriveApp.getFolderById(folderId);
  const folder = DriveApp.createFolder("Fotos privadas dos desafios - Casamento Ana Clara e Felipe");
  properties.setProperty("WEDDING_PRIVATE_CHALLENGE_PHOTOS_FOLDER_ID", folder.getId());
  return folder;
}

function setupChallengeStorage() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("Vincule este script à planilha de confirmações.");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getRegisteredGuestsSheet(spreadsheet);
    getIndividualSubmissionsSheet(spreadsheet);
    const responses = getResponsesSheet(spreadsheet);
    const guestColumn = ensureGuestColumn(responses);
    const headers = responses.getRange(1, 1, 1, responses.getLastColumn()).getDisplayValues()[0];
    const nameColumn = headers.indexOf("Nome") + 1;
    const attendanceColumn = headers.indexOf("Presença") + 1;
    const guestsColumn = guestColumn;
    if (responses.getLastRow() > 1) {
      const rows = responses.getRange(2, 1, responses.getLastRow() - 1, responses.getLastColumn()).getDisplayValues();
      const registry = getRegisteredGuestsSheet(spreadsheet);
      rows.forEach(row => {
        if (row[attendanceColumn - 1] !== "Sim, estarei lá") return;
        const names = [row[nameColumn - 1], ...(row[guestsColumn - 1] || "").split(/\r?\n/)]
          .map(name => name.trim()).filter(Boolean);
        names.forEach(name => ensureRegisteredGuest(registry, name));
      });
    }
    const folder = getPrivateChallengePhotosFolder();
    Logger.log(`Desafios individuais preparados; pasta privada: "${folder.getName()}".`);
  } finally {
    lock.releaseLock();
  }
}
