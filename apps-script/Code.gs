// ============================================================
// AeroStock — backend Google Apps Script
// Telegram : validation par bouton (sondage auto-installé)
// Photos : plusieurs par fiche, stockées sur Google Drive
//
// INSTALLATION (une seule fois) :
//   1. Coller ce fichier dans l'éditeur Apps Script
//   2. Exécuter la fonction setupTelegram()  -> demande token + chat id
//   3. Exécuter la fonction initTelegram()   -> supprime le webhook,
//      installe le déclencheur 1 min, envoie un message de test
//   4. Déployer : Deploy > Manage deployments > modifier > New version
// ============================================================

// ---- Configuration (stockée dans les propriétés du script, jamais dans le code) ----
// À définir dans : Éditeur Apps Script > Paramètres du projet (⚙) > Propriétés du script :
//   TELEGRAM_TOKEN  = token du bot (BotFather)
//   TELEGRAM_CHAT_ID = ex : -5266276172
function getProp_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setupTelegram() {
  // Vérifie la configuration et envoie un message de test (aucune fenêtre : sûr depuis l'éditeur)
  var token = getProp_('TELEGRAM_TOKEN');
  var chat = getProp_('TELEGRAM_CHAT_ID');
  if (!token || !chat) {
    Logger.log(
      'Configuration manquante. Définissez TELEGRAM_TOKEN et TELEGRAM_CHAT_ID dans ' +
      'Paramètres du projet > Propriétés du script, puis relancez setupTelegram().'
    );
    return;
  }
  Logger.log('TELEGRAM_TOKEN : présent (' + token.substring(0, 10) + '…)');
  Logger.log('TELEGRAM_CHAT_ID : ' + chat);
  var res = sendTelegram('Test AeroStock : configuration valide.');
  Logger.log('Message de test envoyé.');
  return res;
}

function initTelegram() {
  deleteTelegramWebhook_();
  var triggerOk = installTrigger_();
  telegramApi_('setMyCommands', {
    commands: JSON.stringify([
      { command: 'valider', description: 'Valider une entree (voir message)' }
    ])
  });
  sendTelegram('Bot AeroStock initialise. Les boutons "Valider" fonctionnent desormais.');
  if (triggerOk) {
    Logger.log('initTelegram OK — webhook supprime, declencheur installe.');
  } else {
    Logger.log(
      'Webhook supprime et bot pret. ATTENTION : le declencheur automatique n a pas pu etre ' +
      'installe (permissions). Creez-le manuellement : icone Horloge (Declencheurs) > ' +
      'Ajouter un declencheur > fonction checkTelegram > Minuteur > Toutes les minutes.'
    );
  }
}

function installTrigger_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'checkTelegram') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    ScriptApp.newTrigger('checkTelegram').timeBased().everyMinutes(1).create();
    return true;
  } catch (e) {
    Logger.log('installTrigger_: ' + e.message);
    return false;
  }
}

function deleteTelegramWebhook_() {
  var res = telegramApi_('deleteWebhook', { drop_pending_updates: 'false' });
  Logger.log('deleteWebhook: ' + JSON.stringify(res));
  return res;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('AeroStock')
      .addItem('Vérifier la config Telegram', 'setupTelegram')
      .addItem('Initialiser (trigger + boutons)', 'initTelegram')
      .addItem('Tester le bot', 'testTelegram')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen (pas de UI): ' + e);
  }
}

function testTelegram() {
  sendTelegram('Test AeroStock : si vous recevez ce message, la configuration est bonne.');
}

function telegramApi_(method, payload) {
  var token = getProp_('TELEGRAM_TOKEN');
  if (!token) throw new Error('TELEGRAM_TOKEN manquant : executez setupTelegram().');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true
  });
  try {
    return JSON.parse(res.getContentText());
  } catch (e) {
    return { ok: false, error: res.getContentText() };
  }
}

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function getLastUpdate() {
  var p = PropertiesService.getScriptProperties();
  return parseInt(p.getProperty('LAST_UPDATE') || '0', 10);
}

function setLastUpdate(id) {
  PropertiesService.getScriptProperties().setProperty('LAST_UPDATE', String(id));
}

function sendTelegram(msg, replyMarkup) {
  var chatId = getProp_('TELEGRAM_CHAT_ID');
  if (!chatId) {
    Logger.log('TELEGRAM_CHAT_ID manquant (Propriétés du script).');
    return { ok: false, error: 'chat_id manquant' };
  }
  var payload = { chat_id: chatId, text: msg };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  try {
    return telegramApi_('sendMessage', payload);
  } catch (e) {
    Logger.log('Telegram send error: ' + e);
    return { ok: false, error: String(e) };
  }
}

function sendNotification(sheetName, action, designation, rowId) {
  var actionLabel = action === 'add' ? '+AJOUT' : action === 'update' ? '~MODIF' : '-SUPPR';
  var msg = '[' + actionLabel + '] ' + sheetName + '\n' + designation + '\n' + new Date().toLocaleString('fr-FR');
  var replyMarkup = null;
  if (action === 'add') {
    msg += '\n\nCliquez sur "Valider" ou repondez : /valider_' + sheetName.replace(/ /g, '_') + ' ' + designation;
    replyMarkup = {
      inline_keyboard: [[{
        text: 'Valider',
        callback_data: 'valider|' + encodeURIComponent(sheetName) + '|' + encodeURIComponent(rowId || '')
      }]]
    };
  }
  sendTelegram(msg, replyMarkup);
}

// Sondage Telegram : lit les clics sur les boutons et les commandes.
// Le declencheur est installe automatiquement par initTelegram().
function checkTelegram() {
  var lastUpdate = getLastUpdate();
  var token = getProp_('TELEGRAM_TOKEN');
  if (!token) return;

  var url = 'https://api.telegram.org/bot' + token + '/getUpdates?offset=' + (lastUpdate + 1) + '&timeout=5';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var data = JSON.parse(resp.getContentText());

    if (!data.ok) {
      // Conflit webhook : on le supprime une fois et on reessaiera au prochain passage.
      if (String(data.description || '').indexOf('webhook') !== -1) {
        Logger.log('Webhook actif detecte, suppression: ' + data.description);
        deleteTelegramWebhook_();
      } else {
        Logger.log('Telegram getUpdates error: ' + resp.getContentText());
      }
      return;
    }

    var updates = data.result || [];
    for (var i = 0; i < updates.length; i++) {
      var upd = updates[i];
      setLastUpdate(upd.update_id);

      if (upd.callback_query) {
        var cbMsg = upd.callback_query.message;
        var cbChatId = cbMsg && cbMsg.chat ? String(cbMsg.chat.id) : '';
        if (cbChatId === getProp_('TELEGRAM_CHAT_ID')) processCallback(upd.callback_query);
        continue;
      }

      if (!upd.message || !upd.message.text) continue;
      var chatId = String(upd.message.chat.id);
      if (chatId !== getProp_('TELEGRAM_CHAT_ID')) continue;
      processCommand(upd.message.text.trim());
    }
  } catch (e) {
    Logger.log('Telegram poll error: ' + e);
  }
}

function processCommand(text) {
  var command = String(text || '').trim();
  var match = command.match(/^\/valider(?:@[\w_]+)?(?:_|\s+)(.+)$/i);
  if (!match) {
    sendTelegram('Format: /valider_Onglet Designation');
    return;
  }

  var args = match[1].trim();
  var separator = args.search(/\s/);
  if (separator < 0) {
    sendTelegram('Format: /valider_Onglet Designation');
    return;
  }

  var sheetName = args.substring(0, separator).replace(/_/g, ' ');
  var designation = args.substring(separator + 1).trim();
  if (!sheetName || !designation) {
    sendTelegram('Format: /valider_Onglet Designation');
    return;
  }

  var result = validateEntry(sheetName, designation, '');
  sendTelegram(result.message);
}

function processCallback(query) {
  var callbackData = String(query.data || '');
  try {
    if (callbackData.toLowerCase().indexOf('valider|') === 0) {
      var parts = callbackData.split('|');
      if (parts.length >= 3) {
        var sheetName = decodeURIComponent(parts[1]);
        var rowId = decodeURIComponent(parts.slice(2).join('|'));
        var result = validateEntry(sheetName, '', rowId);
        if (!updateValidationMessage(query.message, result)) sendTelegram(result.message);
        answerCallback(query.id, result.success ? 'Entree validee' : result.message);
        return;
      }
    }

    if (/^\/valider(?:@[\w_]+)?(?:_|\s)/i.test(callbackData)) {
      processCommand(callbackData);
      answerCallback(query.id, 'Commande traitee');
      return;
    }

    answerCallback(query.id, 'Bouton non reconnu');
  } catch (e) {
    Logger.log('processCallback error: ' + e);
    answerCallback(query.id, 'Erreur: ' + e.message);
  }
}

function updateValidationMessage(message, result) {
  if (!message || !message.chat || !message.message_id) return false;

  var method = result.success ? 'editMessageText' : 'editMessageReplyMarkup';
  var payload = {
    chat_id: message.chat.id,
    message_id: message.message_id,
    reply_markup: JSON.stringify({ inline_keyboard: [] })
  };

  if (result.success) {
    var originalText = String(message.text || '');
    if (originalText.indexOf('Statut : Validé') < 0) {
      originalText += '\n\nStatut : Validé';
    }
    payload.text = originalText;
  } else if (message.text) {
    payload.text = String(message.text);
    method = 'editMessageText';
  }

  try {
    var data = telegramApi_(method, payload);
    return data.ok === true;
  } catch (e) {
    Logger.log('Telegram message update error: ' + e);
    return false;
  }
}

function answerCallback(callbackId, text) {
  try {
    telegramApi_('answerCallbackQuery', {
      callback_query_id: callbackId,
      text: String(text || '').substring(0, 200)
    });
  } catch (e) {
    Logger.log('Telegram callback error: ' + e);
  }
}

function normalizeValue(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/\s+/g, ' ');
}

// Retrouve une feuille par son nom, en ignorant la casse et les espaces
// (permet de renommer un onglet sans casser le site).
function findSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var direct = ss.getSheetByName(name);
  if (direct) return direct;
  var target = normalizeValue(name);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalizeValue(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

function validateEntry(sheetName, designation, rowId) {
  var sheet = findSheet_(sheetName);
  if (!sheet) {
    return { success: false, message: 'Onglet "' + sheetName + '" introuvable.' };
  }

  var expectedId = String(rowId || '').trim();
  var expectedDesignation = normalizeValue(designation);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var matches = expectedId
      ? String(data[i][4] || '').trim() === expectedId
      : normalizeValue(data[i][0]) === expectedDesignation;
    if (!matches) continue;

    var rowDesignation = String(data[i][0] || designation || '').trim();
    if (normalizeValue(data[i][6]) !== 'en attente') {
      return { success: false, message: 'Non trouve ou deja valide: ' + rowDesignation };
    }

    sheet.getRange(i + 1, 7).setValue('Validé');
    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'OK: ' + rowDesignation + ' (' + sheetName + ') validé'
    };
  }

  return { success: false, message: 'Non trouve ou deja valide: ' + (designation || rowId) };
}

// ---------- Photos multiples (Google Drive) ----------

function getPhotoFolder_() {
  var name = 'AeroStock Photos';
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// Convertit une dataURL "data:image/jpeg;base64,...." en fichier Drive public
function saveImageToDrive_(dataUrl, baseName) {
  var match = String(dataUrl || '').match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!match) return String(dataUrl || ''); // deja une URL : on garde
  var mime = match[1];
  var bytes = Utilities.base64Decode(match[2]);
  var blob = Utilities.newBlob(bytes, mime, baseName + '-' + Date.now() + '.jpg');
  var file = getPhotoFolder_().createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('setSharing: ' + e);
  }
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// Entree : chaine JSON (tableau de dataURL ou d'URL) — sortie : URLs jointes par |
function saveImages_(imagesParam, baseName) {
  if (!imagesParam) return '';
  var list;
  try {
    list = JSON.parse(imagesParam);
  } catch (e) {
    list = [imagesParam];
  }
  if (!Array.isArray(list)) list = [list];
  var urls = [];
  for (var i = 0; i < list.length && i < 6; i++) {
    if (!list[i]) continue;
    urls.push(saveImageToDrive_(list[i], baseName || 'photo'));
  }
  return urls.join('|');
}

function handleRequest(e) {
  var sheetName = e.parameter.sheet || 'Feuille 1';
  var sheet = findSheet_(sheetName);
  var action = e.parameter.action;
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  if (!sheet) { return output.setContent(JSON.stringify({ error: 'Feuille introuvable' })); }

  if (!action || action === 'read') {
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        var statut = data[i].length > 6 ? String(data[i][6] || '').trim() : '';
        if (statut === 'En attente') continue;
        var existingId = String(data[i][4] || '').trim();
        if (!existingId) {
          existingId = 'id-' + (i + 1);
          sheet.getRange(i + 1, 5).setValue(existingId);
        }
        result.push({
          designation: String(data[i][0] || ''),
          col2: String(data[i][1] || ''),
          col3: String(data[i][2] || ''),
          col4: String(data[i][3] || ''),
          id: existingId,
          image: data[i].length > 5 ? String(data[i][5] || '') : '',
          col5: data[i].length > 7 ? String(data[i][7] || '') : ''
        });
      }
    }
    return output.setContent(JSON.stringify(result));
  }

  if (action === 'valider') {
    var targetId = e.parameter.id;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4] || '') === targetId && normalizeValue(data[i][6]) === 'en attente') {
        sheet.getRange(i + 1, 7).setValue('Validé');
        return output.setContent(JSON.stringify({ success: true }));
      }
    }
    return output.setContent(JSON.stringify({ error: 'ID introuvable ou deja valide' }));
  }

  if (action === 'pending') {
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && normalizeValue(data[i][6]) === 'en attente') {
        result.push({
          designation: String(data[i][0] || ''),
          col2: String(data[i][1] || ''),
          id: String(data[i][4] || '')
        });
      }
    }
    return output.setContent(JSON.stringify(result));
  }

  if (action === 'add') {
    var rowId = String(Date.now()) + '-' + String(Math.floor(Math.random() * 10000));
    var imagesAdd = e.parameter.images !== undefined
      ? saveImages_(e.parameter.images, e.parameter.designation || 'photo')
      : String(e.parameter.image || '');
    sheet.appendRow([
      e.parameter.designation || '',
      e.parameter.col2 || '',
      e.parameter.col3 || '',
      e.parameter.col4 || '',
      rowId,
      imagesAdd,
      'En attente',
      e.parameter.col5 || ''
    ]);
    sendNotification(sheetName, 'add', e.parameter.designation || '', rowId);
    return output.setContent(JSON.stringify({ success: true, id: rowId, images: imagesAdd }));
  }

  if (action === 'update') {
    var targetId = e.parameter.id;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4] || '') === targetId) {
        sheet.getRange(i + 1, 1).setValue(e.parameter.designation || '');
        sheet.getRange(i + 1, 2).setValue(e.parameter.col2 || '');
        sheet.getRange(i + 1, 3).setValue(e.parameter.col3 || '');
        sheet.getRange(i + 1, 4).setValue(e.parameter.col4 || '');
        if (e.parameter.images !== undefined) {
          sheet.getRange(i + 1, 6).setValue(
            saveImages_(e.parameter.images, e.parameter.designation || 'photo')
          );
        } else if (e.parameter.image !== undefined) {
          sheet.getRange(i + 1, 6).setValue(e.parameter.image || '');
        }
        if (e.parameter.col5 !== undefined) sheet.getRange(i + 1, 8).setValue(e.parameter.col5 || '');
        sendNotification(sheetName, 'update', e.parameter.designation || '');
        return output.setContent(JSON.stringify({ success: true }));
      }
    }
    return output.setContent(JSON.stringify({ error: 'ID introuvable' }));
  }

  if (action === 'delete') {
    var targetId = e.parameter.id;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][4] || '') === targetId) {
        sheet.deleteRow(i + 1);
        sendNotification(sheetName, 'delete', String(data[i][0] || ''));
        return output.setContent(JSON.stringify({ success: true }));
      }
    }
    return output.setContent(JSON.stringify({ error: 'ID introuvable' }));
  }

  return output.setContent(JSON.stringify({ error: 'action inconnue' }));
}