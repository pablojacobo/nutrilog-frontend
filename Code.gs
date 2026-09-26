/** Google Apps Script backend for the meal tracker. */
// Para cambiar la imagen del logo, modificá únicamente el valor de esta variable.
const LOGO_FILE_URL = 'https://drive.google.com/file/d/1Mob4mSLY9Z27LPffYAeuiI7Ia8bNDIYP/view?usp=drive_link';
const CONFIG = {
  SHEET_NAME: 'Registro comidas',
  MEASUREMENTS_SHEET_NAME: 'Registro peso y medidas',
  USERS_SHEET_NAME: 'Usuarios',
  DRIVE_FOLDER_ID: '1mxSQ-qtCT2m8UxKcaJ5JcywRbDEzhmz_',
  GROQ_API_URL: 'https://api.groq.com/openai/v1/chat/completions',
  GROQ_MODEL: 'meta-llama/llama-4-maverick-17b-128e-instruct',
  GEMINI_API_URL: 'https://generativelanguage.googleapis.com/v1beta/models/',
  GEMINI_MODEL: 'gemini-3.6-flash',
  HEADERS: ['Día', 'Hora', 'Comida', 'Calorías estimadas', 'Proteínas estimadas', 'H.C. estimados (g)', 'Grasas estimadas (g)', 'IMG de la comida', 'Comentario/Evento', 'Usuario'],
  MEASUREMENT_HEADERS: ['Día', 'Hora', 'Peso (kg)', 'Grasa abdominal (%)', 'Grasa visceral', 'Músculo (%)', 'Usuario'],
  USER_HEADERS: ['Usuario', 'Nombre', 'Contraseña', 'Rol', 'Alta', 'Nutricionista', 'Mail nutricionista', 'Fecha de nacimiento', 'Estatura (cm)', 'Actividad', 'Calorías mínimas', 'Calorías máximas', 'Proteínas mínimas (g/día)', 'Proteínas máximas (g/día)', 'H.C. mínimos (g/día)', 'H.C. máximos (g/día)', 'Grasas mínimas (g/día)', 'Grasas máximas (g/día)'],
  ACTIVITY_OPTIONS: [
    'Sedentario (poco o ningún ejercicio). (vida de oficina, sin ejercicio programado).',
    'Ligeramente activo (ejercicio ligero 1 a 3 días por semana).',
    'Moderadamente activo (ejercicio moderado 3 a 5 días por semana)',
    'Muy activo (ejercicio fuerte 6 a 7 días por semana).',
    'Extremadamente activo (entrenamientos muy duros o doble sesión diaria).'
  ]
};

function doGet() {
  ensureSheet_();
  const template = HtmlService.createTemplateFromFile('Index');
  template.logoUrl = getLogoUrl_();
  return template.evaluate()
    .setTitle('NutriLog. Tu alimentación, en un solo lugar')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Endpoint POST para frontend estático (GitHub Pages). El frontend envía { action, payload }.
function doPost(e) {
  try {
    const body = JSON.parse(e.postData || '{}');
    const action = body.action;
    const payload = body.payload || {};
    const fn = this[action];
    if (typeof fn !== 'function') {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Acción no encontrada: ' + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const result = fn(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message || String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getLogoImage() {
  return getLogoUrl_();
}

function setupApp() {
  ensureSheet_();
  ensureMeasurementsSheet_();
  ensureUsersSheet_();
  return 'Hojas listas: ' + CONFIG.SHEET_NAME + ', ' + CONFIG.MEASUREMENTS_SHEET_NAME + ' y ' + CONFIG.USERS_SHEET_NAME;
}

function createWeeklyReportTrigger() {
  ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === 'sendWeeklyReports';
  }).forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });
  ScriptApp.newTrigger('sendWeeklyReports').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  return 'Activador semanal creado para los lunes a las 08:00.';
}

function clearLegacySessions() {
  const properties = PropertiesService.getScriptProperties();
  const all = properties.getProperties();
  Object.keys(all).filter(function(key) {
    return key.indexOf('session:') === 0;
  }).forEach(function(key) { properties.deleteProperty(key); });
  return 'Sesiones antiguas eliminadas.';
}

function saveMeal(payload) {
  const user = requireSession_(payload && payload.sessionToken);
  if (!payload || !payload.date || !payload.time || (!payload.description && !(payload.image && payload.image.base64))) {
    throw new Error('Indicá una descripción o adjuntá una foto de la comida.');
  }
  const saveId = String(payload.saveId || '').trim();
  if (!saveId) throw new Error('No se pudo identificar el intento de guardado. Volvé a intentarlo.');
  const cache = CacheService.getScriptCache();
  const cacheKey = 'meal-save:' + user.username + ':' + saveId;
  const cachedResult = cache.get(cacheKey);
  if (cachedResult) {
    const parsedResult = JSON.parse(cachedResult);
    if (parsedResult && parsedResult.calories !== undefined && parsedResult.protein !== undefined && parsedResult.carbs !== undefined && parsedResult.fat !== undefined) {
      return parsedResult;
    }
    cache.remove(cacheKey);
  }

  const estimate = analyzeMeal_(payload.description, payload.image);
  const savedMeal = payload.description ? String(payload.description) : 'Comida identificada en la foto';
  let imageUrl = '';
  if (payload.image && payload.image.base64) {
    imageUrl = saveImage_(payload.image, payload.date, payload.time);
  }

  const sheet = ensureSheet_();
  sheet.appendRow([
    payload.date,
    payload.time,
    savedMeal,
    estimate.calories,
    estimate.protein,
    estimate.carbs,
    estimate.fat,
    imageUrl,
    payload.comment || '',
    user.username
  ]);

  const result = {
    meal: savedMeal,
    calories: estimate.calories,
    protein: estimate.protein,
    carbs: estimate.carbs,
    fat: estimate.fat,
    imageUrl: imageUrl
  };
  cache.put(cacheKey, JSON.stringify(result), 21600);
  return result;
}

function analyzeMealPreview(payload) {
  const user = requireSession_(payload && payload.sessionToken);
  if (!payload || (!payload.description && !(payload.image && payload.image.base64))) {
    throw new Error('Indicá una descripción o adjuntá una foto de la comida.');
  }
  let temporaryFileName = '';
  if (payload.image && payload.image.base64) {
    temporaryFileName = saveTemporaryMealImage_(payload.image, getPreviewOwnerKey_(user));
  }
  try {
    const estimate = analyzeMeal_(payload.description, payload.image);
    return {
      meal: payload.description ? String(payload.description) : 'Comida identificada en la foto',
      calories: estimate.calories,
      protein: estimate.protein,
      carbs: estimate.carbs,
      fat: estimate.fat
    };
  } finally {
    if (temporaryFileName) deleteTemporaryMealImage_(temporaryFileName);
  }
}

function describeMeal(payload) {
  const image = payload && payload.image ? payload.image : payload;
  if (!image || !image.base64) throw new Error('No se recibió una imagen.');

  const apiKeys = getApiKeys_();
  if (!apiKeys.length) {
    throw new Error('Faltan claves Groq. Configurá GROQ_API_KEY o GROQ_API_KEYS_JSON en las propiedades del proyecto.');
  }

  const prompt = [
    'Identificá únicamente los alimentos visibles dentro del plato o recipiente principal.',
    'Ignorá la mesa, cubiertos, envases, decoración y el fondo. No describas la presentación.',
    'Escribí en español argentino (rioplatense), con vocabulario natural y voseo cuando corresponda.',
    'Respondé con una lista breve en una sola frase, sin cantidades, calorías, proteínas ni explicaciones.',
    'Respondé ÚNICAMENTE JSON válido con este formato exacto:',
    '{"meal":"descripción breve de la comida"}',
    'No inventes alimentos que no se vean claramente. No inventes campos.'
  ].join('\n');

  const shuffled = apiKeys.slice().sort(function() { return Math.random() - 0.5; });
  let lastError = '';
  const attemptedModels = [];
  const attempts = Math.max(shuffled.length * 2, 2);
  for (let i = 0; i < attempts; i++) {
    try {
      if (i === shuffled.length) attemptedModels.length = 0;
      const apiKey = shuffled[i % shuffled.length];
      const endpoint = CONFIG.GROQ_API_URL;
      const model = getGroqModel_(true, attemptedModels, undefined, true);
      attemptedModels.push(model);
      const body = {
        model: model,
        messages: [{ role: 'user', content: [
          { type: 'text', text: prompt },
          { type: 'image_url', image_url: { url: 'data:' + image.mimeType + ';base64,' + image.base64 } }
        ] }],
        temperature: 0,
        max_tokens: 120,
        response_format: { type: 'json_object' }
      };
      const response = UrlFetchApp.fetch(endpoint, {
        method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey }, payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      const status = response.getResponseCode();
      if (status < 200 || status >= 300) {
        if (status === 429 || status === 500 || status === 502 || status === 503 || status === 504) {
          Utilities.sleep(700 * (i + 1));
        }
        throw new Error('HTTP ' + status + ': ' + response.getContentText().slice(0, 300));
      }
      const raw = JSON.parse(response.getContentText());
      const result = parseGroqJson_(raw.choices[0].message.content);
      if (!result.meal) throw new Error('Respuesta descriptiva incompleta');
      return String(result.meal).slice(0, 500);
    } catch (error) {
      lastError = error.message;
    }
  }
  throw new Error('Servidor ocupado, por favor, intentá nuevamente más tarde.');
}

function getMeals(filters) {
  const user = requireSession_(filters && filters.sessionToken);
  const sheet = ensureSheet_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return [];

  const from = filters && filters.from ? filters.from : '0000-00-00';
  const to = filters && filters.to ? filters.to : '9999-12-31';
  const patientUsername = getTargetPatient_(user, filters && filters.patientUsername);
  return values.slice(1)
    .map(function(row, index) { return { row: index + 2, values: row }; })
    .filter(function(item) { return item.values[0] >= from && item.values[0] <= to && item.values[9] === patientUsername; })
    .map(function(item) {
      const row = item.values;
      return {
        row: item.row,
        date: row[0], time: row[1], meal: row[2], calories: Number(row[3]) || 0,
        protein: Number(row[4]) || 0, carbs: Number(row[5]) || 0, fat: Number(row[6]) || 0,
        imageUrl: row[7], comment: row[8]
      };
    })
    .sort(function(a, b) {
      return compareDateTime_(a.date, a.time, b.date, b.time);
    });
}

function deleteMeal(payload) {
  const user = requireSession_(payload && payload.sessionToken);
  const rowNumber = Number(payload && payload.row);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) throw new Error('No se pudo identificar el registro.');
  const sheet = ensureSheet_();
  const row = sheet.getRange(rowNumber, 1, 1, CONFIG.HEADERS.length).getDisplayValues()[0];
  if (!row[9]) throw new Error('Ese registro ya no existe.');
  const targetUsername = getTargetPatient_(user, payload && payload.patientUsername);
  if (row[9] !== targetUsername) throw new Error('No tenés permisos para borrar ese registro.');
  sheet.deleteRow(rowNumber);
  return true;
}

function getCurrentUser(sessionToken) {
  return requireSession_(sessionToken);
}

function registerUser(payload) {
  throw new Error('Los usuarios son dados de alta por un nutricionista.');
}

function loginUser(payload) {
  const rawUsername = String(payload && payload.username || '');
  if (!rawUsername || /\s/.test(rawUsername)) throw new Error('El nombre de usuario no puede contener espacios.');
  const username = normalizeUsername_(rawUsername);
  const password = String(payload && payload.password || '');
  const user = findUser_(ensureUsersSheet_(), username);
  if (!user || user.passwordHash !== hashPassword_(password)) throw new Error('Nombre de usuario o contraseña incorrectos.');
  const accessType = String(payload && payload.accessType || 'usuario').toLowerCase();
  if (accessType === 'nutricionista' && user.role !== 'nutricionista') throw new Error('Esta cuenta no tiene acceso de nutricionista.');
  if (accessType === 'usuario' && user.role !== 'usuario') throw new Error('Ingresá desde el acceso de nutricionistas.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session:' + token, JSON.stringify({ username: user.username, name: user.name, role: user.role }), 21600);
  return { sessionToken: token, username: user.username, name: user.name, role: user.role };
}

function logoutUser(payload) {
  const sessionToken = (payload && payload.sessionToken) || payload;
  if (sessionToken) CacheService.getScriptCache().remove('session:' + sessionToken);
  return true;
}

function changePassword(payload) {
  const current = requireSession_(payload && payload.sessionToken);
  const password = String(payload && payload.password || '');
  if (password.length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres.');
  const sheet = ensureUsersSheet_();
  const user = findUser_(sheet, current.username);
  sheet.getRange(user.row, 3).setValue(hashPassword_(password));
  return true;
}

function getPatients(sessionToken) {
  const nutritionist = requireRole_(sessionToken, 'nutricionista');
  return getUsers_().filter(function(user) { return user.role === 'usuario' && user.nutritionist === nutritionist.username; }).map(function(user) {
    return { username: user.username, name: user.name, profile: profileFromUser_(user) };
  });
}

function getPatientProfile(payload) {
  const user = requireSession_(payload && payload.sessionToken);
  const targetUsername = getTargetPatient_(user, payload && payload.patientUsername);
  const patient = getUsers_().filter(function(item) { return item.username === targetUsername; })[0];
  return profileFromUser_(patient);
}

function savePatientProfile(payload) {
  const user = requireSession_(payload && payload.sessionToken);
  const targetUsername = getTargetPatient_(user, payload && payload.patientUsername);
  const patient = getUsers_().filter(function(item) { return item.username === targetUsername; })[0];
  const sheet = ensureUsersSheet_();
  const birthDate = String(payload.birthDate || '').trim();
  const height = payload.height === '' || payload.height === null || payload.height === undefined ? '' : Number(payload.height);
  const minCalories = payload.minCalories === '' || payload.minCalories === null || payload.minCalories === undefined ? '' : Number(payload.minCalories);
  const maxCalories = payload.maxCalories === '' || payload.maxCalories === null || payload.maxCalories === undefined ? '' : Number(payload.maxCalories);
  const minProtein = payload.minProtein === '' || payload.minProtein === null || payload.minProtein === undefined ? '' : Number(payload.minProtein);
  const maxProtein = payload.maxProtein === '' || payload.maxProtein === null || payload.maxProtein === undefined ? '' : Number(payload.maxProtein);
  const minCarbs = payload.minCarbs === '' || payload.minCarbs === null || payload.minCarbs === undefined ? '' : Number(payload.minCarbs);
  const maxCarbs = payload.maxCarbs === '' || payload.maxCarbs === null || payload.maxCarbs === undefined ? '' : Number(payload.maxCarbs);
  const minFat = payload.minFat === '' || payload.minFat === null || payload.minFat === undefined ? '' : Number(payload.minFat);
  const maxFat = payload.maxFat === '' || payload.maxFat === null || payload.maxFat === undefined ? '' : Number(payload.maxFat);
  if (birthDate && !/^\d{4}-\d{2}-\d{2}$/.test(birthDate)) throw new Error('Ingresá una fecha de nacimiento válida.');
  if (height !== '' && (!Number.isFinite(height) || height <= 0)) throw new Error('Ingresá una estatura válida.');
  if (minCalories !== '' && (!Number.isFinite(minCalories) || minCalories < 0)) throw new Error('Ingresá un mínimo calórico válido.');
  if (maxCalories !== '' && (!Number.isFinite(maxCalories) || maxCalories < 0 || (minCalories !== '' && maxCalories < minCalories))) throw new Error('El máximo calórico debe ser mayor o igual al mínimo.');
  const macroRanges = [
    { min: minProtein, max: maxProtein, name: 'proteínas' },
    { min: minCarbs, max: maxCarbs, name: 'hidratos de carbono' },
    { min: minFat, max: maxFat, name: 'grasas' }
  ];
  macroRanges.forEach(function(m) {
    if (m.min !== '' && (!Number.isFinite(m.min) || m.min < 0)) throw new Error('Ingresá un mínimo de ' + m.name + ' válido.');
    if (m.max !== '' && (!Number.isFinite(m.max) || m.max < 0)) throw new Error('Ingresá un máximo de ' + m.name + ' válido.');
    if (m.min !== '' && m.max !== '' && m.max < m.min) throw new Error('El máximo de ' + m.name + ' debe ser mayor o igual al mínimo.');
  });
  sheet.getRange(patient.row, 7, 1, 12).setValues([[
    patient.nutritionistEmail || getNutritionistEmail_(patient.nutritionist), birthDate, height,
    normalizeActivity_(payload.activity), minCalories, maxCalories,
    payload.minProtein === '' || payload.minProtein === null || payload.minProtein === undefined ? '' : Number(payload.minProtein),
    payload.maxProtein === '' || payload.maxProtein === null || payload.maxProtein === undefined ? '' : Number(payload.maxProtein),
    payload.minCarbs === '' || payload.minCarbs === null || payload.minCarbs === undefined ? '' : Number(payload.minCarbs),
    payload.maxCarbs === '' || payload.maxCarbs === null || payload.maxCarbs === undefined ? '' : Number(payload.maxCarbs),
    payload.minFat === '' || payload.minFat === null || payload.minFat === undefined ? '' : Number(payload.minFat),
    payload.maxFat === '' || payload.maxFat === null || payload.maxFat === undefined ? '' : Number(payload.maxFat)
  ]]);
  return getPatientProfile({ sessionToken: payload.sessionToken, patientUsername: targetUsername });
}

function resetPatientPassword(payload) {
  const nutritionist = requireRole_(payload && payload.sessionToken, 'nutricionista');
  const patientUsername = normalizeUsername_(payload && payload.patientUsername);
  const password = String(payload && payload.password || '');
  if (password.length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres.');
  const patient = getUsers_().filter(function(user) {
    return user.username === patientUsername && user.role === 'usuario' && user.nutritionist === nutritionist.username;
  })[0];
  if (!patient) throw new Error('Ese usuario no pertenece a tu lista.');
  ensureUsersSheet_().getRange(patient.row, 3).setValue(hashPassword_(password));
  return true;
}

function createPatient(payload) {
  const nutritionist = requireRole_(payload && payload.sessionToken, 'nutricionista');
  const rawUsername = String(payload && payload.username || '');
  const username = normalizeUsername_(rawUsername);
  const name = String(payload && payload.name || '').trim();
  const password = String(payload && payload.password || '');
  if (!rawUsername || /\s/.test(rawUsername)) throw new Error('El nombre de usuario no puede contener espacios.');
  if (!name || password.length < 6) throw new Error('Completá nombre, usuario y una contraseña de al menos 6 caracteres.');
  const sheet = ensureUsersSheet_();
  if (findUser_(sheet, username)) throw new Error('Ya existe ese nombre de usuario.');
  sheet.appendRow([username, name, hashPassword_(password), 'Usuario', new Date(), nutritionist.username, getNutritionistEmail_(nutritionist.username), '', '', '', '', '', '', '', '', '', '', '']);
  return { username: username, name: name };
}

function saveMeasurement(payload) {
  const user = requireSession_(payload && payload.sessionToken);
  if (!payload || !payload.date || !payload.time || payload.weight === '' || payload.weight === null || payload.weight === undefined) {
    throw new Error('El peso, el día y la hora son obligatorios.');
  }
  const targetUsername = getTargetPatient_(user, payload && payload.patientUsername);
  const values = [payload.date, payload.time, Number(payload.weight), optionalNumber_(payload.abdominalFat), optionalNumber_(payload.visceralFat), optionalNumber_(payload.muscle), targetUsername];
  if (!Number.isFinite(values[2]) || values[2] <= 0) throw new Error('Ingresá un peso válido.');
  ensureMeasurementsSheet_().appendRow(values);
  return true;
}

function getMeasurements(filters) {
  const user = requireSession_(filters && filters.sessionToken);
  const patientUsername = getTargetPatient_(user, filters && filters.patientUsername);
  const from = filters && filters.from ? filters.from : '0000-00-00';
  const to = filters && filters.to ? filters.to : '9999-12-31';
  return ensureMeasurementsSheet_().getDataRange().getDisplayValues().slice(1)
    .filter(function(row) { return row[0] >= from && row[0] <= to && row[6] === patientUsername; })
    .map(function(row) { return { date: row[0], time: row[1], weight: row[2], abdominalFat: row[3], visceralFat: row[4], muscle: row[5] }; })
    .sort(function(a, b) { return compareDateTime_(a.date, a.time, b.date, b.time); });
}

function analyzeMeal_(description, image) {
  const apiKeys = getApiKeys_();
  if (!apiKeys.length) {
    throw new Error('Faltan claves Groq. Configurá GROQ_API_KEY o GROQ_API_KEYS_JSON en las propiedades del proyecto.');
  }

  const prompt = [
    'Estimá rápidamente calorías, proteínas, hidratos de carbono y grasas de una porción razonable.',
    'Usá la descripción y la imagen si existe. Considerá solo los alimentos dentro del plato.',
    'No expliques el cálculo. Devolvé solo cuatro campos y números enteros.',
    'Respondé ÚNICAMENTE JSON válido con este formato exacto:',
    '{"calories":0,"protein":0,"carbs":0,"fat":0}',
    'calories, protein, carbs y fat son números enteros sin unidades. No inventes campos.',
    'Descripción del usuario: ' + (description || '(sin descripción; identificar la comida en la imagen)')
  ].join('\n');

  const shuffled = apiKeys.slice().sort(function() { return Math.random() - 0.5; });
  let lastError = '';
  const attemptedModels = [];
  const attempts = Math.max(shuffled.length * 2, 2);
  for (let i = 0; i < attempts; i++) {
    try {
      if (i === shuffled.length) attemptedModels.length = 0;
      const apiKey = shuffled[i % shuffled.length];
      const endpoint = CONFIG.GROQ_API_URL;
      const model = getGroqModel_(!!(image && image.base64), attemptedModels);
      attemptedModels.push(model);
      const body = {
        model: model,
        messages: [{ role: 'user', content: [{ type: 'text', text: prompt }] }],
        temperature: 0,
        max_tokens: 120
      };
      if (image && image.base64) {
        body.messages[0].content.push({ type: 'image_url', image_url: { url: 'data:' + image.mimeType + ';base64,' + image.base64 } });
      }
      const response = UrlFetchApp.fetch(endpoint, {
        method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + apiKey }, payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      const status = response.getResponseCode();
      if (status < 200 || status >= 300) {
        if (status === 429 || status === 500 || status === 502 || status === 503 || status === 504) {
          Utilities.sleep(500 * (i + 1));
        }
        let apiMessage = response.getContentText();
        try {
          const apiError = JSON.parse(apiMessage);
          apiMessage = apiError.error && apiError.error.message ? apiError.error.message : apiMessage;
        } catch (parseError) {
          // Conserva el texto original cuando la API no devuelve JSON.
        }
        throw new Error('HTTP ' + status + ': ' + apiMessage.slice(0, 300));
      }
      const raw = JSON.parse(response.getContentText());
      const result = parseGroqJson_(raw.choices[0].message.content);
      if (!Number.isFinite(Number(result.calories)) || !Number.isFinite(Number(result.protein)) || !Number.isFinite(Number(result.carbs)) || !Number.isFinite(Number(result.fat))) {
        throw new Error('Respuesta nutricional incompleta');
      }
      return {
        calories: Math.max(0, Math.round(Number(result.calories))),
        protein: Math.max(0, Math.round(Number(result.protein))),
        carbs: Math.max(0, Math.round(Number(result.carbs))),
        fat: Math.max(0, Math.round(Number(result.fat)))
      };
    } catch (error) {
      lastError = error.message;
    }
  }
  throw new Error('Servidor ocupado, por favor, intentá nuevamente más tarde.');
}

function saveImage_(image, date, time) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const extension = (image.mimeType || 'image/jpeg').split('/')[1] || 'jpg';
  const safeName = ('comida_' + date + '_' + time).replace(/[^a-zA-Z0-9_-]/g, '-') + '.' + extension;
  const blob = Utilities.newBlob(Utilities.base64Decode(image.base64), image.mimeType, safeName);
  return folder.createFile(blob).getUrl();
}

function parseGroqJson_(content) {
  const text = String(content || '').replace(/```json|```/gi, '').trim();
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('Groq no devolvió un objeto JSON válido.');
  return JSON.parse(match[0]);
}

function saveTemporaryMealImage_(image, username) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const fileName = 'vista_previa_comida_' + username + '_' + Utilities.getUuid() + '.jpg';
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(image.base64), image.mimeType || 'image/jpeg', fileName);
    folder.createFile(blob);
    return fileName;
  } finally {
    lock.releaseLock();
  }
}

function deleteTemporaryMealImage_(fileName) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) files.next().setTrashed(true);
}

function getPreviewOwnerKey_(user) {
  const patient = getUsers_().filter(function(item) { return item.username === user.username; })[0];
  const nutritionist = user.role === 'nutricionista' ? user.username : (patient && patient.nutritionist) || 'sin_nutricionista';
  return (nutritionist + '_' + user.username).replace(/[^a-zA-Z0-9_-]/g, '-');
}

function getApiKeys_() {
  const properties = PropertiesService.getScriptProperties();
  const singleKey = String(properties.getProperty('GROQ_API_KEY') || '').trim();
  const raw = properties.getProperty('GROQ_API_KEYS_JSON');
  if (singleKey) return [singleKey];
  if (!raw || !raw.trim()) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter(function(key) { return typeof key === 'string' && key.trim(); }) : [];
  } catch (error) {
    throw new Error('GROQ_API_KEYS_JSON no contiene JSON válido.');
  }
}

function getGeminiApiKeys_() {
  const raw = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEYS_JSON');
  if (!raw || !raw.trim()) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter(function(key) {
      return typeof key === 'string' && key.trim();
    }).map(function(key) { return key.trim(); }) : [];
  } catch (error) {
    throw new Error('GEMINI_API_KEYS_JSON no contiene un array JSON válido.');
  }
}

function getGeminiModel_() {
  const configuredModel = String(PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || CONFIG.GEMINI_MODEL).trim();
  return /^gemini-2\.5-flash(?:-|$)/i.test(configuredModel) ? CONFIG.GEMINI_MODEL : configuredModel;
}

function getGeminiModelCandidates_(apiKey) {
  const configuredModel = getGeminiModel_();
  const cache = CacheService.getScriptCache();
  const keyFingerprint = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, apiKey, Utilities.Charset.UTF_8)
    .map(function(byte) { return (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'); }).join('');
  const cacheKey = 'gemini-model-candidates:' + keyFingerprint;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed) && parsed.length && parsed.every(function(model) { return typeof model === 'string' && model; })) {
        return parsed;
      }
    } catch (error) {}
  }

  const endpoint = CONFIG.GEMINI_API_URL + '?key=' + encodeURIComponent(apiKey);
  try {
    const response = UrlFetchApp.fetch(endpoint, { method: 'get', muteHttpExceptions: true, timeout: 12000 });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return [configuredModel];
    const raw = JSON.parse(response.getContentText());
    const available = (raw.models || []).filter(function(model) {
      return model && model.name && (!model.supportedGenerationMethods || model.supportedGenerationMethods.indexOf('generateContent') !== -1);
    }).map(function(model) {
      return String(model.name).replace(/^models\//, '');
    }).filter(function(model) {
      return /gemini/i.test(model) && !/embedding|aqa|robotics|tts|image-generation/i.test(model);
    });
    const unique = available.filter(function(model, index) { return available.indexOf(model) === index; });
    const stable = unique.filter(function(model) {
      const normalized = model.toLowerCase();
      return normalized.indexOf('latest') === -1 && normalized.indexOf('preview') === -1 &&
        normalized.indexOf('thinking') === -1 && normalized.indexOf('experimental') === -1;
    });
    const preferred = stable.filter(function(model) {
      return /^gemini-\d+(?:\.\d+)*-flash(?:-|$)/i.test(model);
    }).sort(compareGeminiModelVersions_);
    const remaining = stable.filter(function(model) { return preferred.indexOf(model) === -1; });
    const ordered = [];
    const addModel = function(model) {
      if (model && ordered.indexOf(model) === -1) ordered.push(model);
    };
    addModel(configuredModel);
    preferred.forEach(addModel);
    remaining.forEach(addModel);
    unique.forEach(addModel);
    if (!ordered.length) ordered.push(configuredModel);
    const result = ordered.slice(0, 4);
    cache.put(cacheKey, JSON.stringify(result), 1800);
    return result;
  } catch (error) {
    console.error('No se pudieron consultar los modelos Gemini: ' + (error.message || error));
    return [configuredModel];
  }
}

function compareGeminiModelVersions_(left, right) {
  const leftVersion = String(left).match(/^gemini-(\d+(?:\.\d+)*)/i);
  const rightVersion = String(right).match(/^gemini-(\d+(?:\.\d+)*)/i);
  if (leftVersion && rightVersion) {
    const leftParts = leftVersion[1].split('.').map(Number);
    const rightParts = rightVersion[1].split('.').map(Number);
    const parts = Math.max(leftParts.length, rightParts.length);
    for (let index = 0; index < parts; index++) {
      const leftPart = leftParts[index] || 0;
      const rightPart = rightParts[index] || 0;
      if (leftPart !== rightPart) return rightPart - leftPart;
    }
  }
  const leftPriority = /-lite(?:-|$)/i.test(left) ? 1 : 0;
  const rightPriority = /-lite(?:-|$)/i.test(right) ? 1 : 0;
  if (leftPriority !== rightPriority) return leftPriority - rightPriority;
  return String(left).localeCompare(String(right));
}

function testGeminiApiKeys() {
  const apiKeys = getGeminiApiKeys_();
  if (!apiKeys.length) throw new Error('No hay claves Gemini configuradas.');
  return apiKeys.map(function(apiKey, index) {
    const model = getGeminiModelCandidates_(apiKey)[0];
    const endpoint = CONFIG.GEMINI_API_URL + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKey);
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'Respondé únicamente OK.' }] }], generationConfig: { maxOutputTokens: 10 } }),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    let detail = response.getContentText().slice(0, 300);
    try {
      const parsed = JSON.parse(response.getContentText());
      detail = status >= 200 && status < 300 && parsed.candidates
        ? String(parsed.candidates[0].content.parts[0].text || '').trim()
        : (parsed.error && parsed.error.message) || detail;
    } catch (error) {}
    return { keyNumber: index + 1, model: model, status: status, ok: status >= 200 && status < 300, detail: detail };
  });
}

function getGroqModel_(requiresVision, excludedModels, apiKey, useConfiguredDirectly) {
  excludedModels = excludedModels || [];
  const configuredModel = String(PropertiesService.getScriptProperties().getProperty('GROQ_MODEL') || CONFIG.GROQ_MODEL).trim();
  const availableApiKeys = apiKey ? [apiKey] : getApiKeys_();
  if (!availableApiKeys.length) return configuredModel || CONFIG.GROQ_MODEL;

  const cache = CacheService.getScriptCache();
  const cacheSuffix = apiKey ? ':' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, apiKey, Utilities.Charset.UTF_8).map(function(byte) {
    return (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0');
  }).join('') : '';
  const cacheKey = 'groq-models:v3:' + (requiresVision ? 'vision' : 'text') + cacheSuffix;
  const response = UrlFetchApp.fetch('https://api.groq.com/openai/v1/models', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + availableApiKeys[0] },
    muteHttpExceptions: true, timeout: 12000
  });
  const status = response.getResponseCode();
  let modelRecords = [];
  if (status >= 200 && status < 300) {
    try {
      const raw = JSON.parse(response.getContentText());
      modelRecords = (raw.data || []).filter(function(model) {
        return model && model.id && model.active !== false;
      });
    } catch (error) {}
  } else if (useConfiguredDirectly !== false && configuredModel && excludedModels.indexOf(configuredModel) === -1 && (!requiresVision || isGroqVisionModel_(configuredModel))) {
    return configuredModel;
  } else {
    throw new Error('No se pudieron consultar los modelos disponibles.');
  }

  const modelIds = modelRecords.map(function(model) { return String(model.id); });
  const candidates = getGroqModelCandidates_(modelRecords, requiresVision)
    .filter(function(model) { return excludedModels.indexOf(model) === -1; });
  const cachedModel = cache.get(cacheKey);
  if (cachedModel && modelIds.indexOf(cachedModel) !== -1 && excludedModels.indexOf(cachedModel) === -1) {
    return cachedModel;
  }
  if (candidates.length) {
    const selectedModel = candidates[0];
    cache.put(cacheKey, selectedModel, 300);
    return selectedModel;
  }
  if (useConfiguredDirectly !== false && configuredModel && excludedModels.indexOf(configuredModel) === -1 && isGroqChatModel_(configuredModel, requiresVision)) {
    return configuredModel;
  }
  const available = modelIds.filter(function(model) {
    return excludedModels.indexOf(model) === -1 && isGroqChatModel_(model, requiresVision);
  }).slice(0, 12).join(', ');
  throw new Error('No hay un modelo ' + (requiresVision ? 'multimodal' : 'de texto') + ' disponible para esta clave. Modelos informados: ' + (available || 'ninguno') + '.');
}

function getGroqModelCandidates_(modelRecords, requiresVision) {
  const candidates = modelRecords
    .filter(function(model) { return isGroqChatModel_(model, requiresVision); })
    .map(function(model) { return String(model.id); });
  const sorted = candidates.sort(compareGroqModelPriority_);
  return sorted;
}

function compareGroqModelPriority_(left, right) {
  const leftPriority = getGroqModelPriority_(left);
  const rightPriority = getGroqModelPriority_(right);
  return leftPriority - rightPriority || String(left).localeCompare(String(right));
}

function getGroqModelPriority_(model) {
  const modelId = String(model || '').toLowerCase();
  if (/^qwen\/qwen3\.(?:6|8)-27b$/i.test(modelId)) return 0;
  if (/^openai\/gpt-oss-20b$/i.test(modelId)) return 1;
  if (/^openai\/gpt-oss-120b$/i.test(modelId)) return 2;
  if (/llama-3\.3-70b-versatile|llama-3\.1-8b-instant/i.test(modelId)) return 3;
  if (/^meta-llama\/llama/i.test(modelId) || /^qwen\//i.test(modelId) || /^google\/gemma/i.test(modelId)) return 4;
  return 5;
}

function isGroqChatModel_(model, requiresVision) {
  const record = typeof model === 'string' ? { id: model } : (model || {});
  const modelId = String(record.id || '').toLowerCase();
  if (!modelId || /whisper|orpheus|prompt-guard|guard|embedding|audio|speech|tts|stt|asr|transcription|safeguard|safety/i.test(modelId)) {
    return false;
  }
  if (requiresVision) return isGroqVisionModel_(record);
  return true;
}

function isGroqVisionModel_(model) {
  const record = typeof model === 'string' ? { id: model } : (model || {});
  const modelId = String(record.id || '').toLowerCase();
  const modalities = record.input_modalities || (record.architecture && record.architecture.input_modalities);
  if (Array.isArray(modalities) && modalities.some(function(modality) { return String(modality).toLowerCase() === 'image'; })) {
    return true;
  }
  return /vision|vl|maverick|scout|llama-4|qwen.*vl/i.test(modelId);
}

function listGroqModels() {
  const apiKeys = getApiKeys_();
  if (!apiKeys.length) throw new Error('No hay claves Groq configuradas.');
  const response = UrlFetchApp.fetch('https://api.groq.com/openai/v1/models', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + apiKeys[0] },
    muteHttpExceptions: true
  });
  return response.getContentText();
}

function testQwenGroqModels() {
  const apiKeys = getApiKeys_();
  if (!apiKeys.length) throw new Error('No hay claves Groq configuradas.');

  const models = ['qwen/qwen3.6-27b', 'qwen/qwen3.8-27b'];
  return models.map(function(model) {
    const response = UrlFetchApp.fetch(CONFIG.GROQ_API_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKeys[0] },
      payload: JSON.stringify({
        model: model,
        messages: [{ role: 'user', content: 'Respondé únicamente: OK' }],
        temperature: 0,
        max_tokens: 10
      }),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    const body = response.getContentText();
    let detail = body.slice(0, 300);
    try {
      const parsed = JSON.parse(body);
      detail = status >= 200 && status < 300
        ? parsed.choices[0].message.content
        : (parsed.error && parsed.error.message) || detail;
    } catch (error) {
      // Conserva el texto original cuando la API no devuelve JSON.
    }
    return { model: model, status: status, ok: status >= 200 && status < 300, detail: detail };
  });
}

function ensureSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  const legacyHeaders = ['Día', 'Hora', 'Comida', 'Calorías estimadas', 'Proteínas estimadas', 'IMG de la comida', 'Comentario/Evento', 'Usuario'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold');
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.HEADERS.length)).getValues()[0];
    if (currentHeaders.slice(0, legacyHeaders.length).join('|') === legacyHeaders.join('|')) {
      sheet.insertColumns(6, 2);
    }
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureMeasurementsSheet_() {
  return ensureStructuredSheet_(CONFIG.MEASUREMENTS_SHEET_NAME, CONFIG.MEASUREMENT_HEADERS);
}

function ensureUsersSheet_() {
  const sheet = ensureStructuredSheet_(CONFIG.USERS_SHEET_NAME, CONFIG.USER_HEADERS);
  if (sheet.getLastRow() > 1) {
    const roles = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
    const normalizedRoles = roles.map(function(row) {
      return [String(row[0]).toLowerCase() === 'nutricionista' ? 'Nutricionista' : 'Usuario'];
    });
    sheet.getRange(2, 4, normalizedRoles.length, 1).setValues(normalizedRoles);
  }
  return sheet;
}

function getLogoUrl_() {
  const driveIdMatch = LOGO_FILE_URL.match(/(?:\/d\/|[?&]id=|^)([a-zA-Z0-9_-]{20,})(?:[/?&]|$)/);
  if (driveIdMatch) {
    const thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + driveIdMatch[1] + '&sz=w1408';
    try {
      const response = UrlFetchApp.fetch(thumbnailUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        const blob = response.getBlob();
        return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
      }
    } catch (error) {
      // Usa la URL de Drive como respaldo si la descarga del servidor falla.
    }
    return thumbnailUrl;
  }
  return LOGO_FILE_URL;
}

function ensureStructuredSheet_(name, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else if (sheet.getRange(1, 1, 1, headers.length).getValues()[0].join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  // Ensure sheet has enough columns for data
  const lastCol = sheet.getLastColumn();
  if (lastCol < headers.length) {
    sheet.insertColumnsAfter(lastCol, headers.length - lastCol);
  }
  return sheet;
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeUsername_(value) {
  return String(value || '').trim().toLowerCase();
}

function compareDateTime_(dateA, timeA, dateB, timeB) {
  const dateKeyA = dateSortKey_(dateA);
  const dateKeyB = dateSortKey_(dateB);
  if (dateKeyA !== dateKeyB) return dateKeyA.localeCompare(dateKeyB);
  return timeSortKey_(timeA) - timeSortKey_(timeB);
}

function dateSortKey_(value) {
  const text = String(value || '').trim();
  let match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (match) return match[1] + padSortPart_(match[2]) + padSortPart_(match[3]);
  match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (match) return match[3] + padSortPart_(match[2]) + padSortPart_(match[1]);
  return text;
}

function timeSortKey_(value) {
  const match = String(value || '').trim().match(/^(\d{1,2}):?(\d{2})/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : Number.MAX_SAFE_INTEGER;
}

function padSortPart_(value) {
  return String(value).padStart(2, '0');
}

function hashPassword_(password) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8)
    .map(function(byte) { return (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'); }).join('');
}

function optionalNumber_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new Error('Las medidas opcionales deben ser números positivos.');
  return number;
}

function getUsers_() {
  const values = ensureUsersSheet_().getDataRange().getValues();
  return values.slice(1).filter(function(row) { return row[0]; }).map(function(row, index) {
    const role = String(row[3]).toLowerCase() === 'nutricionista' ? 'nutricionista' : 'usuario';
    return {
      row: index + 2, username: normalizeUsername_(row[0]), name: String(row[1]), passwordHash: String(row[2]),
      role: role, nutritionist: normalizeUsername_(row[5]), nutritionistEmail: normalizeEmail_(row[6]),
      birthDate: formatSheetDate_(row[7]), height: row[8] === '' || row[8] === undefined ? '' : Number(row[8]), activity: normalizeActivity_(row[9]),
      minCalories: row[10] === '' || row[10] === undefined ? '' : Number(row[10]), maxCalories: row[11] === '' || row[11] === undefined ? '' : Number(row[11]),
      minProtein: row[12] === '' || row[12] === undefined ? '' : Number(row[12]), maxProtein: row[13] === '' || row[13] === undefined ? '' : Number(row[13]),
      minCarbs: row[14] === '' || row[14] === undefined ? '' : Number(row[14]), maxCarbs: row[15] === '' || row[15] === undefined ? '' : Number(row[15]),
      minFat: row[16] === '' || row[16] === undefined ? '' : Number(row[16]), maxFat: row[17] === '' || row[17] === undefined ? '' : Number(row[17])
    };
  });
}

function profileFromUser_(user) {
  return {
    username: user.username, name: user.name, birthDate: user.birthDate || '', height: user.height || '',
    activity: user.activity || '', minCalories: user.minCalories || '', maxCalories: user.maxCalories || '',
    minProtein: user.minProtein || '', maxProtein: user.maxProtein || '',
    minCarbs: user.minCarbs || '', maxCarbs: user.maxCarbs || '',
    minFat: user.minFat || '', maxFat: user.maxFat || ''
  };
}

function normalizeActivity_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const normalized = text.toLowerCase();
  if (/sedent/.test(normalized)) return CONFIG.ACTIVITY_OPTIONS[0];
  if (/lig(er|e)amente activo|1 a 3/.test(normalized)) return CONFIG.ACTIVITY_OPTIONS[1];
  if (/moderad/.test(normalized) || /3 a 5/.test(normalized)) return CONFIG.ACTIVITY_OPTIONS[2];
  if (/muy activo|6 a 7/.test(normalized)) return CONFIG.ACTIVITY_OPTIONS[3];
  if (/extremad|doble sesi/.test(normalized)) return CONFIG.ACTIVITY_OPTIONS[4];
  return text;
}

function formatSheetDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  return '';
}

function getNutritionistEmail_(username) {
  const nutritionist = getUsers_().filter(function(user) { return user.username === normalizeUsername_(username); })[0];
  return nutritionist ? nutritionist.nutritionistEmail : '';
}

function sendWeeklyReports() {
  const today = new Date();
  const endDate = new Date(today.getTime() - 24 * 60 * 60 * 1000);
    const startDate = new Date(endDate.getTime() - 6 * 24 * 60 * 60 * 1000);
  const timeZone = Session.getScriptTimeZone();
  const from = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
  const to = Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd');
  const patients = getUsers_().filter(function(user) { return user.role === 'usuario' && user.nutritionist; });
  const groups = {};
  patients.forEach(function(patient) {
    (groups[patient.nutritionist] = groups[patient.nutritionist] || []).push(patient);
  });
  Object.keys(groups).forEach(function(username) {
    const nutritionist = getUsers_().filter(function(user) { return user.username === username; })[0];
    const email = nutritionist && nutritionist.nutritionistEmail;
    if (!email) return;
    const sections = groups[username].map(function(patient) {
      try {
        return buildWeeklyPatientReport_(patient, from, to);
      } catch (error) {
        console.error('No se pudo armar el reporte semanal para ' + patient.username + ': ' + (error.message || error));
        return 'PACIENTE: ' + patient.name + '\nNo se pudo generar el reporte de este paciente. Revisá sus registros y datos del perfil.';
      }
    });
    if (!sections.length) return;
      const body = 'Resumen semanal NutriLog\nPeríodo: ' + from + ' a ' + to + '\n\n' + sections.join('\n\n');
    MailApp.sendEmail({ to: email, subject: 'Resumen semanal de tus pacientes - NutriLog', body: body });
  });
}

function sendWeeklyReportsGroq() {
  const today = new Date();
  const endDate = new Date(today.getTime() - 24 * 60 * 60 * 1000);
  const startDate = new Date(endDate.getTime() - 6 * 24 * 60 * 60 * 1000);
  const timeZone = Session.getScriptTimeZone();
  const from = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
  const to = Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd');
  const patients = getUsers_().filter(function(user) { return user.role === 'usuario' && user.nutritionist; });
  const groups = {};
  patients.forEach(function(patient) {
    (groups[patient.nutritionist] = groups[patient.nutritionist] || []).push(patient);
  });
  Object.keys(groups).forEach(function(username) {
    const nutritionist = getUsers_().filter(function(user) { return user.username === username; })[0];
    const email = nutritionist && nutritionist.nutritionistEmail;
    if (!email) return;
    const sections = groups[username].map(function(patient) {
      try {
        return buildWeeklyPatientReportGroq_(patient, from, to);
      } catch (error) {
        console.error('No se pudo armar el reporte semanal para ' + patient.username + ': ' + (error.message || error));
        return 'PACIENTE: ' + patient.name + '\nNo se pudo generar el reporte de este paciente. Revisá sus registros y datos del perfil.';
      }
    });
    if (!sections.length) return;
    const body = 'Resumen semanal NutriLog\nPeríodo: ' + from + ' a ' + to + '\n\n' + sections.join('\n\n');
    MailApp.sendEmail({ to: email, subject: 'Resumen semanal de tus pacientes - NutriLog', body: body });
  });
}

function buildWeeklyPatientReportGroq_(patient, from, to) {
  const meals = getMealsForReport_(patient.username, from, to);
  const measurement = getLatestMeasurementForReport_(patient.username, to);
  const totals = meals.reduce(function(result, meal) {
    result.calories += meal.calories; result.protein += meal.protein; result.carbs += meal.carbs; result.fat += meal.fat; return result;
  }, { calories: 0, protein: 0, carbs: 0, fat: 0 });
  const recommendation = generateWeeklyRecommendationGroq_(patient, measurement, meals, totals, from, to);
  const stats = getWeeklyReportStats_(meals);
  const dataSummary = [
    'Datos registrados:',
    '- Comidas: ' + meals.length + ' en ' + stats.days + ' días.',
    '- Calorías: ' + totals.calories + ' kcal totales; promedio de ' + stats.averageCalories + ' kcal por día con registros.',
    '- Proteínas: ' + totals.protein + ' g totales; promedio de ' + stats.averageProtein + ' g por día con registros.',
    '- H.C.: ' + totals.carbs + ' g totales; promedio de ' + stats.averageCarbs + ' g por día con registros.',
    '- Grasas: ' + totals.fat + ' g totales; promedio de ' + stats.averageFat + ' g por día con registros.',
    '- Edad: ' + (patient.birthDate ? calculateAge_(patient.birthDate, new Date()) + ' años' : 'no informada') + '.',
    '- Estatura: ' + (patient.height || 'no informada') + (patient.height ? ' cm.' : '.'),
    '- Actividad: ' + (patient.activity || 'no informada') + '.',
    '- Rango calórico objetivo: ' + (patient.minCalories !== '' ? patient.minCalories : 'no informado') + ' a ' + (patient.maxCalories !== '' ? patient.maxCalories : 'no informado') + ' kcal/día.',
    '- Objetivo de proteínas: ' + (patient.minProtein !== '' || patient.maxProtein !== '' ? (patient.minProtein || 'sin mínimo') + ' a ' + (patient.maxProtein || 'sin máximo') + ' g/día' : 'no informado') + '.',
    '- Objetivo de H.C.: ' + (patient.minCarbs !== '' || patient.maxCarbs !== '' ? (patient.minCarbs || 'sin mínimo') + ' a ' + (patient.maxCarbs || 'sin máximo') + ' g/día' : 'no informado') + '.',
    '- Objetivo de grasas: ' + (patient.minFat !== '' || patient.maxFat !== '' ? (patient.minFat || 'sin mínimo') + ' a ' + (patient.maxFat || 'sin máximo') + ' g/día' : 'no informado') + '.',
    '- Última medición: ' + (measurement ? 'peso ' + measurement.weight + ' kg; grasa abdominal ' + (measurement.abdominalFat || 'no informada') + '; grasa visceral ' + (measurement.visceralFat || 'no informada') + '; músculo ' + (measurement.muscle || 'no informado') + '.' : 'no informada.')
  ].join('\n');
  return 'PACIENTE: ' + patient.name + '\n' + dataSummary + '\n\nRecomendación profesional:\n' + recommendation;
}

function buildWeeklyPatientReport_(patient, from, to) {
  const meals = getMealsForReport_(patient.username, from, to);
  const measurement = getLatestMeasurementForReport_(patient.username, to);
  const totals = meals.reduce(function(result, meal) {
    result.calories += meal.calories; result.protein += meal.protein; result.carbs += meal.carbs; result.fat += meal.fat; return result;
  }, { calories: 0, protein: 0, carbs: 0, fat: 0 });
  const recommendation = generateWeeklyRecommendation_(patient, measurement, meals, totals, from, to);
  const stats = getWeeklyReportStats_(meals);
  const dataSummary = [
    'Datos registrados:',
    '- Comidas: ' + meals.length + ' en ' + stats.days + ' días.',
    '- Calorías: ' + totals.calories + ' kcal totales; promedio de ' + stats.averageCalories + ' kcal por día con registros.',
    '- Proteínas: ' + totals.protein + ' g totales; promedio de ' + stats.averageProtein + ' g por día con registros.',
    '- H.C.: ' + totals.carbs + ' g totales; promedio de ' + stats.averageCarbs + ' g por día con registros.',
    '- Grasas: ' + totals.fat + ' g totales; promedio de ' + stats.averageFat + ' g por día con registros.',
    '- Edad: ' + (patient.birthDate ? calculateAge_(patient.birthDate, new Date()) + ' años' : 'no informada') + '.',
    '- Estatura: ' + (patient.height || 'no informada') + (patient.height ? ' cm.' : '.'),
    '- Actividad: ' + (patient.activity || 'no informada') + '.',
    '- Rango calórico objetivo: ' + (patient.minCalories !== '' ? patient.minCalories : 'no informado') + ' a ' + (patient.maxCalories !== '' ? patient.maxCalories : 'no informado') + ' kcal/día.',
    '- Objetivo de proteínas: ' + (patient.minProtein !== '' || patient.maxProtein !== '' ? (patient.minProtein || 'sin mínimo') + ' a ' + (patient.maxProtein || 'sin máximo') + ' g/día' : 'no informado') + '.',
    '- Objetivo de H.C.: ' + (patient.minCarbs !== '' || patient.maxCarbs !== '' ? (patient.minCarbs || 'sin mínimo') + ' a ' + (patient.maxCarbs || 'sin máximo') + ' g/día' : 'no informado') + '.',
    '- Objetivo de grasas: ' + (patient.minFat !== '' || patient.maxFat !== '' ? (patient.minFat || 'sin mínimo') + ' a ' + (patient.maxFat || 'sin máximo') + ' g/día' : 'no informado') + '.',
    '- Última medición: ' + (measurement ? 'peso ' + measurement.weight + ' kg; grasa abdominal ' + (measurement.abdominalFat || 'no informada') + '; grasa visceral ' + (measurement.visceralFat || 'no informada') + '; músculo ' + (measurement.muscle || 'no informado') + '.' : 'no informada.')
  ].join('\n');
  return 'PACIENTE: ' + patient.name + '\n' + dataSummary + '\n\nRecomendación profesional:\n' + recommendation;
}

function getMealsForReport_(username, from, to) {
  return ensureSheet_().getDataRange().getDisplayValues().slice(1).filter(function(row) {
    return row[9] === username && row[0] >= from && row[0] <= to;
  }).map(function(row) {
    return { date: row[0], meal: row[2], calories: Number(row[3]) || 0, protein: Number(row[4]) || 0, carbs: Number(row[5]) || 0, fat: Number(row[6]) || 0 };
  });
}

function getLatestMeasurementForReport_(username, to) {
  const rows = ensureMeasurementsSheet_().getDataRange().getDisplayValues().slice(1).filter(function(row) {
    return row[6] === username && row[0] <= to;
  }).sort(function(a, b) { return compareDateTime_(a[0], a[1], b[0], b[1]); });
  if (!rows.length) return null;
  return { date: rows[rows.length - 1][0], weight: rows[rows.length - 1][2], abdominalFat: rows[rows.length - 1][3], visceralFat: rows[rows.length - 1][4], muscle: rows[rows.length - 1][5] };
}

function generateWeeklyRecommendation_(patient, measurement, meals, totals, from, to) {
  let apiKeys;
  try {
    apiKeys = getGeminiApiKeys_();
  } catch (error) {
    console.warn('No se pudieron leer las claves Gemini para el reporte semanal: ' + (error.message || error));
    return buildSafeWeeklyFallback_(patient, measurement, meals, totals);
  }
  if (!apiKeys.length) return buildSafeWeeklyFallback_(patient, measurement, meals, totals);
  const prompt = buildWeeklyRecommendationPrompt_(patient, measurement, meals, totals, from, to);
  const startedAt = new Date().getTime();
  const maxPatientMs = 20000;
  const maxAttempts = 4;
  let lastError = 'Error desconocido';
  let stopReason = '';
  let attempt = 0;
  let consecutiveRetryableFailures = 0;
  let discoveryFailures = 0;
  let stopTrying = false;
  for (let index = 0; index < apiKeys.length && !stopTrying; index++) {
    if (new Date().getTime() - startedAt > maxPatientMs) {
      stopReason = 'se agotó el tiempo máximo por paciente';
      break;
    }
    let models;
    try {
      models = getGeminiModelCandidates_(apiKeys[index]);
    } catch (error) {
      lastError = error.message || String(error);
      discoveryFailures++;
      consecutiveRetryableFailures++;
      console.warn('Reporte semanal, clave Gemini ' + (index + 1) + ': no se pudieron obtener modelos: ' + lastError);
      if (discoveryFailures >= 2 || consecutiveRetryableFailures >= 2 || new Date().getTime() - startedAt > maxPatientMs) {
        stopReason = 'falló la consulta de modelos';
        stopTrying = true;
      }
      continue;
    }
    const modelsToTry = models.slice(0, 2);
    for (let modelIndex = 0; modelIndex < modelsToTry.length && !stopTrying; modelIndex++) {
      if (attempt >= maxAttempts || new Date().getTime() - startedAt > maxPatientMs) {
        stopReason = attempt >= maxAttempts ? 'se alcanzó el máximo de intentos' : 'se agotó el tiempo máximo por paciente';
        stopTrying = true;
        break;
      }
      const model = modelsToTry[modelIndex];
      attempt++;
      try {
        const endpoint = CONFIG.GEMINI_API_URL + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(apiKeys[index]);
        return requestWeeklyGeminiRecommendation_(endpoint, model, prompt);
      } catch (error) {
        lastError = error.message || String(error);
        console.warn('Reporte semanal, clave Gemini ' + (index + 1) + ', modelo ' + model + ': ' + lastError);
        const retryable = isRetryableWeeklyGeminiError_(lastError);
        if (retryable) {
          consecutiveRetryableFailures++;
        } else {
          consecutiveRetryableFailures = 0;
        }
        if (retryable && consecutiveRetryableFailures >= 2) {
          stopReason = 'Gemini respondió de forma repetida que está ocupado o incompleto';
          stopTrying = true;
          break;
        }
        if (attempt >= maxAttempts || new Date().getTime() - startedAt > maxPatientMs) {
          stopReason = attempt >= maxAttempts ? 'se alcanzó el máximo de intentos' : 'se agotó el tiempo máximo por paciente';
          stopTrying = true;
          break;
        }
        if (retryable) {
          Utilities.sleep(Math.min(1000, 250 * attempt));
        }
      }
    }
  }
  if (!stopReason) stopReason = 'no hubo una respuesta utilizable';
  console.warn('Se usará el respaldo del reporte semanal para ' + patient.username + ': ' + lastError + ' (' + stopReason + ').');
  return buildSafeWeeklyFallback_(patient, measurement, meals, totals);
}

function generateWeeklyRecommendationGroq_(patient, measurement, meals, totals, from, to) {
  let apiKeys;
  try {
    apiKeys = getApiKeys_();
  } catch (error) {
    console.warn('No se pudieron leer las claves del servicio para el reporte semanal: ' + (error.message || error));
    return buildSafeWeeklyFallback_(patient, measurement, meals, totals);
  }
  if (!apiKeys.length) return buildSafeWeeklyFallback_(patient, measurement, meals, totals);
  const prompt = buildGroqWeeklyRecommendationPrompt_(patient, measurement, meals, totals, from, to);
  const startedAt = new Date().getTime();
  const maxPatientMs = 30000;
  const maxAttempts = 6;
  let lastError = 'Error desconocido';
  let stopReason = '';
  let attempt = 0;
  let stopTrying = false;
  for (let index = 0; index < apiKeys.length && !stopTrying; index++) {
    const attemptedModels = [];
    if (new Date().getTime() - startedAt > maxPatientMs) {
      stopReason = 'se agotó el tiempo máximo por paciente';
      break;
    }
    for (let modelAttempt = 0; modelAttempt < 3 && !stopTrying; modelAttempt++) {
      if (attempt >= maxAttempts || new Date().getTime() - startedAt > maxPatientMs) {
        stopReason = attempt >= maxAttempts ? 'se alcanzó el máximo de intentos' : 'se agotó el tiempo máximo por paciente';
        stopTrying = true;
        break;
      }
      let model;
      try {
        model = getGroqModel_(false, attemptedModels, apiKeys[index], false);
      } catch (error) {
        lastError = error.message || String(error);
        console.warn('Reporte semanal, clave ' + (index + 1) + ': no se pudo seleccionar modelo: ' + lastError);
        break;
      }
      if (attemptedModels.indexOf(model) !== -1) {
        lastError = 'El servicio devolvió un modelo que ya fue intentado.';
        console.warn('Reporte semanal, clave ' + (index + 1) + ': ' + lastError);
        break;
      }
      attemptedModels.push(model);
      attempt++;
      try {
        return requestGroqWeeklyRecommendation_(apiKeys[index], model, prompt);
      } catch (error) {
        lastError = error.message || String(error);
        console.warn('Reporte semanal, clave ' + (index + 1) + ', modelo ' + model + ': ' + lastError);
        if (attempt >= maxAttempts || new Date().getTime() - startedAt > maxPatientMs) {
          stopReason = attempt >= maxAttempts ? 'se alcanzó el máximo de intentos' : 'se agotó el tiempo máximo por paciente';
          stopTrying = true;
          break;
        }
        if (isRetryableGroqError_(lastError)) {
          Utilities.sleep(Math.min(1000, 250 * attempt));
        }
      }
    }
  }
  if (!stopReason) stopReason = 'no hubo una respuesta utilizable';
  console.warn('Se usará el respaldo del reporte semanal para ' + patient.username + ': ' + lastError + ' (' + stopReason + ').');
  return buildSafeWeeklyFallback_(patient, measurement, meals, totals);
}

function buildGroqWeeklyRecommendationPrompt_(patient, measurement, meals, totals, from, to) {
  const stats = getWeeklyReportStats_(meals);
  const missing = [];
  if (!patient.birthDate) missing.push('fecha de nacimiento');
  if (!patient.height) missing.push('estatura');
  if (!patient.activity) missing.push('nivel de actividad');
  if (!measurement) missing.push('peso y composición corporal');
  const age = patient.birthDate ? calculateAge_(patient.birthDate, new Date()) : 'no informada';
  const objectiveText = [
    'calorías: ' + (patient.minCalories !== '' && patient.maxCalories !== '' ? patient.minCalories + ' a ' + patient.maxCalories + ' kcal/día' : 'no informado'),
    'proteínas: ' + (patient.minProtein !== '' || patient.maxProtein !== '' ? (patient.minProtein || 'sin mínimo') + ' a ' + (patient.maxProtein || 'sin máximo') + ' g/día' : 'no informado'),
    'H.C.: ' + (patient.minCarbs !== '' || patient.maxCarbs !== '' ? (patient.minCarbs || 'sin mínimo') + ' a ' + (patient.maxCarbs || 'sin máximo') + ' g/día' : 'no informado'),
    'grasas: ' + (patient.minFat !== '' || patient.maxFat !== '' ? (patient.minFat || 'sin mínimo') + ' a ' + (patient.maxFat || 'sin máximo') + ' g/día' : 'no informado')
  ].join('; ');
  const mealText = meals.slice(0, 8).map(function(meal) {
    return meal.date + ': ' + String(meal.meal || 'sin descripción').slice(0, 50) + ' (' + meal.calories + ' kcal, ' + meal.protein + ' g de proteínas, ' + meal.carbs + ' g de H.C., ' + meal.fat + ' g de grasas)';
  }).join(' | ');
  const omittedMeals = meals.length > 8 ? ' Se omitieron registros repetidos para mantener breve el contexto.' : '';
  return [
    'Analizá de forma prudente y realista el resumen semanal. Respondé exactamente con estas tres líneas, sin Markdown ni texto adicional:',
    'Fortalezas: una oración clara y accionable.',
    'Ajustes prioritarios: una oración clara y accionable.',
    'Próximo paso: una oración clara y accionable.',
    'No diagnostiques, no inventes alimentos, mediciones, déficits ni objetivos. Usá sólo los datos del contexto.',
    'Compará los promedios diarios con los objetivos informados sólo si están disponibles; aclará que corresponden a los días con registros.',
    'Si no hay comidas, indicá que no es posible evaluar la ingesta y pedí comenzar el registro.',
    'Mencioná datos faltantes en el próximo paso. Máximo 60 palabras en total y terminá cada línea con un punto.',
    'Contexto:',
    'Período: ' + from + ' a ' + to,
    'Paciente: ' + patient.name + '; edad: ' + age + '; estatura: ' + (patient.height || 'no informada') + ' cm; actividad: ' + (patient.activity || 'no informada') + '.',
    'Objetivo: ' + objectiveText + '.',
    'Registros: ' + meals.length + ' comidas en ' + stats.days + ' de 7 días; promedio diario de ' + stats.averageCalories + ' kcal, ' + stats.averageProtein + ' g de proteínas, ' + stats.averageCarbs + ' g de H.C. y ' + stats.averageFat + ' g de grasas por día con registros.',
    (mealText || 'No hay comidas registradas.') + omittedMeals,
    'Totales: ' + totals.calories + ' kcal, ' + totals.protein + ' g de proteínas, ' + totals.carbs + ' g de H.C. y ' + totals.fat + ' g de grasas.',
    'Última medición: ' + formatMeasurementForPrompt_(measurement) + '.',
    'Datos faltantes: ' + (missing.length ? missing.join(', ') : 'ninguno') + '.'
  ].join('\n');
}

function requestGroqWeeklyRecommendation_(apiKey, model, prompt) {
  const payload = {
    model: model,
    messages: [
      { role: 'system', content: 'Sos un nutricionista experto. Respondé en español argentino y respetá las tres partes solicitadas.' },
      { role: 'user', content: prompt }
    ],
    temperature: 0.1,
    max_completion_tokens: 512
  };
  return requestGroqWeeklyRecommendationPayload_(apiKey, payload, model);
}

function requestGroqWeeklyRecommendationPayload_(apiKey, payload, model) {
  const response = UrlFetchApp.fetch(CONFIG.GROQ_API_URL, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 12000
  });
  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status < 200 || status >= 300) {
    let detail = body.slice(0, 300);
    try {
      const parsedError = JSON.parse(body);
      detail = parsedError.error && parsedError.error.message ? parsedError.error.message : detail;
    } catch (parseError) {}
    const error = new Error('HTTP ' + status + ': ' + detail);
    if (status === 400) error.retryWithText = true;
    throw error;
  }
  const raw = JSON.parse(body);
  const choice = raw.choices && raw.choices[0] ? raw.choices[0] : {};
  const content = choice.message && choice.message.content ? String(choice.message.content).trim() : '';
  if (!content) {
    throw new Error('El servicio no devolvió contenido con el modelo ' + model + '. finish_reason: ' + (choice.finish_reason || 'desconocido') + '.');
  }
  const finishReason = String(choice.finish_reason || '').toLowerCase();
  const completeReasons = ['stop', 'eos', 'end_turn', 'complete'];
  if (finishReason && completeReasons.indexOf(finishReason) === -1) {
    throw new Error('Groq no completó la recomendación con el modelo ' + model + '. finish_reason: ' + choice.finish_reason + '.');
  }
  return parseWeeklyRecommendationContent_(content);
}

function isRetryableGroqError_(message) {
  return /^(HTTP (429|500|502|503|504)\b)/i.test(message || '') || /finish_reason:\s*length/i.test(message || '');
}

function buildWeeklyRecommendationPrompt_(patient, measurement, meals, totals, from, to) {
  const stats = getWeeklyReportStats_(meals);
  const missing = [];
  if (!patient.birthDate) missing.push('fecha de nacimiento');
  if (!patient.height) missing.push('estatura');
  if (!patient.activity) missing.push('nivel de actividad');
  if (!measurement) missing.push('peso y composición corporal');
  const age = patient.birthDate ? calculateAge_(patient.birthDate, new Date()) : 'no informada';
  const objectiveText = [
    'calorías: ' + (patient.minCalories !== '' && patient.maxCalories !== '' ? patient.minCalories + ' a ' + patient.maxCalories + ' kcal/día' : 'no informado'),
    'proteínas: ' + (patient.minProtein !== '' || patient.maxProtein !== '' ? (patient.minProtein || 'sin mínimo') + ' a ' + (patient.maxProtein || 'sin máximo') + ' g/día' : 'no informado'),
    'H.C.: ' + (patient.minCarbs !== '' || patient.maxCarbs !== '' ? (patient.minCarbs || 'sin mínimo') + ' a ' + (patient.maxCarbs || 'sin máximo') + ' g/día' : 'no informado'),
    'grasas: ' + (patient.minFat !== '' || patient.maxFat !== '' ? (patient.minFat || 'sin mínimo') + ' a ' + (patient.maxFat || 'sin máximo') + ' g/día' : 'no informado')
  ].join('; ');
  const mealText = meals.slice(0, 12).map(function(meal) {
    return meal.date + ': ' + String(meal.meal || 'sin descripción').slice(0, 80) + ' (' + meal.calories + ' kcal, ' + meal.protein + ' g de proteínas, ' + meal.carbs + ' g de H.C., ' + meal.fat + ' g de grasas)';
  }).join(' | ');
  const omittedMeals = meals.length > 12 ? ' Se omitieron ' + (meals.length - 12) + ' registros repetidos para mantener breve el contexto.' : '';
  return [
    'Comportate como un nutricionista experto. Analizá de forma prudente y realista el resumen semanal de un paciente.',
    'Respondé ÚNICAMENTE JSON válido, sin Markdown ni texto adicional, con este esquema exacto:',
    '{"fortalezas":"...","ajustesPrioritarios":"...","proximoPaso":"..."}',
    'Cada valor debe ser una oración completa en español argentino, clara y accionable. Usá exactamente esos tres campos.',
    'No diagnostiques, no inventes alimentos, mediciones, déficits ni objetivos. Usá sólo los datos del contexto.',
    'Compará los promedios diarios con los objetivos informados sólo si están disponibles; aclará que corresponden a los días con registros.',
    'Si no hay comidas registradas, indicá que no es posible evaluar la ingesta y pedí comenzar el registro.',
    'Mencioná datos faltantes en el próximo paso. Máximo 70 palabras en total y terminá cada campo con un punto.',
    'Contexto:',
    'Período: ' + from + ' a ' + to,
    'Paciente: ' + patient.name + '; edad: ' + age + '; estatura: ' + (patient.height || 'no informada') + ' cm; actividad: ' + (patient.activity || 'no informada') + '.',
    'Objetivo: ' + objectiveText + '.',
    'Última medición: ' + formatMeasurementForPrompt_(measurement),
    'Registros: ' + meals.length + ' comidas en ' + stats.days + ' de 7 días; promedio diario de ' + stats.averageCalories + ' kcal, ' + stats.averageProtein + ' g de proteínas, ' + stats.averageCarbs + ' g de H.C. y ' + stats.averageFat + ' g de grasas por día con registros.',
    mealText + omittedMeals,
    'Totales del período: ' + totals.calories + ' kcal, ' + totals.protein + ' g de proteínas, ' + totals.carbs + ' g de H.C. y ' + totals.fat + ' g de grasas.',
    'Datos faltantes: ' + (missing.length ? missing.join(', ') : 'ninguno') + '.'
  ].join('\n');
}

function requestWeeklyGeminiRecommendation_(endpoint, model, prompt) {
  const structuredPayload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 220,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          fortalezas: { type: 'STRING' },
          ajustesPrioritarios: { type: 'STRING' },
          proximoPaso: { type: 'STRING' }
        },
        required: ['fortalezas', 'ajustesPrioritarios', 'proximoPaso']
      }
    }
  };
  try {
    return requestWeeklyGeminiRecommendationPayload_(endpoint, structuredPayload, model);
  } catch (error) {
    if (!error.retryWithText) throw error;
    const textPayload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 220 }
    };
    return requestWeeklyGeminiRecommendationPayload_(endpoint, textPayload, model);
  }
}

function requestWeeklyGeminiRecommendationPayload_(endpoint, payload, model) {
  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true, timeout: 12000
  });
  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status < 200 || status >= 300) {
    let detail = body.slice(0, 300);
    try {
      const parsedError = JSON.parse(body);
      detail = parsedError.error && parsedError.error.message ? parsedError.error.message : detail;
    } catch (parseError) {}
    const error = new Error('HTTP ' + status + ' con Gemini (' + model + '): ' + detail);
    if (status === 400) error.retryWithText = true;
    throw error;
  }
  const raw = JSON.parse(body);
  const candidate = raw.candidates && raw.candidates[0] ? raw.candidates[0] : {};
  const content = candidate.content && candidate.content.parts
    ? candidate.content.parts.map(function(part) { return part.text || ''; }).join('').trim()
    : '';
  if (!content) {
    const finishReason = candidate.finishReason || 'desconocido';
    const promptFeedback = raw.promptFeedback ? JSON.stringify(raw.promptFeedback) : 'sin detalles de promptFeedback';
    throw new Error('Gemini no devolvió contenido con el modelo ' + model + '. finishReason: ' + finishReason + '; ' + promptFeedback);
  }
  if (candidate.finishReason && candidate.finishReason !== 'STOP') {
    throw new Error('Gemini no completó la recomendación con el modelo ' + model + '. finishReason: ' + candidate.finishReason + '.');
  }
  return parseWeeklyRecommendationContent_(content);
}

function parseWeeklyRecommendationContent_(content) {
  const text = String(content || '').replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/i, '').trim();
  let parsed = null;
  try {
    parsed = JSON.parse(text);
  } catch (jsonError) {
    const match = text.match(/\{[\s\S]*\}/);
    if (match) {
      try { parsed = JSON.parse(match[0]); } catch (nestedError) {}
    }
  }
  if (parsed && typeof parsed === 'object') {
    const fortalezas = String(parsed.fortalezas || '').trim();
    const ajustesPrioritarios = String(parsed.ajustesPrioritarios || '').trim();
    const proximoPaso = String(parsed.proximoPaso || '').trim();
    if (fortalezas && ajustesPrioritarios && proximoPaso) {
      return formatWeeklyRecommendationSections_({
        fortalezas: fortalezas,
        ajustesPrioritarios: ajustesPrioritarios,
        proximoPaso: proximoPaso
      });
    }
  }
  const normalized = normalizePlainWeeklyRecommendation_(text);
  if (!normalized) throw new Error('Gemini devolvió una recomendación incompleta o con formato inesperado.');
  return normalized;
}

function normalizePlainWeeklyRecommendation_(text) {
  if (!/fortalezas\s*:/i.test(text) || !/ajustes\s+prioritarios\s*:/i.test(text) || !/próximo\s+paso\s*:/i.test(text)) return null;
  const lines = text.split(/\n+/).map(function(line) { return line.trim(); }).filter(Boolean);
  if (lines.length < 3) return null;
  const normalized = lines.slice(0, 3).map(function(line) {
    return line.replace(/^(fortalezas|ajustes\s+prioritarios|próximo\s+paso)\s*:/i, function(match) {
      return match.charAt(0).toUpperCase() + match.slice(1);
    }).trim();
  });
  if (!normalized.every(function(line) { return /[.!?]$/.test(line); })) return null;
  return normalized.join('\n');
}

function formatWeeklyRecommendationSections_(sections) {
  const values = ['fortalezas', 'ajustesPrioritarios', 'proximoPaso'].map(function(key) {
    const value = String(sections[key] || '').replace(/\s+/g, ' ').trim();
    return value && /[.!?]$/.test(value) ? value : value + '.';
  });
  if (values.some(function(value) { return !value || value === '.'; })) {
    throw new Error('La recomendación no contiene las tres partes requeridas.');
  }
  const result = 'Fortalezas: ' + values[0] + '\nAjustes prioritarios: ' + values[1] + '\nPróximo paso: ' + values[2];
  if (!isCompleteRecommendation_(result)) throw new Error('La recomendación no superó la validación de formato.');
  return result;
}

function getWeeklyReportStats_(meals) {
  const days = meals.filter(function(meal, index, all) {
    return all.findIndex(function(item) { return item.date === meal.date; }) === index;
  }).length;
  return {
    days: days,
    averageCalories: days ? Math.round(totalsCalories_(meals) / days) : 0,
    averageProtein: days ? Math.round(totalsProtein_(meals) / days) : 0,
    averageCarbs: days ? Math.round(totalsCarbs_(meals) / days) : 0,
    averageFat: days ? Math.round(totalsFat_(meals) / days) : 0
  };
}

function totalsCalories_(meals) {
  return meals.reduce(function(total, meal) { return total + (Number(meal.calories) || 0); }, 0);
}

function totalsProtein_(meals) {
  return meals.reduce(function(total, meal) { return total + (Number(meal.protein) || 0); }, 0);
}

function totalsCarbs_(meals) {
  return meals.reduce(function(total, meal) { return total + (Number(meal.carbs) || 0); }, 0);
}

function totalsFat_(meals) {
  return meals.reduce(function(total, meal) { return total + (Number(meal.fat) || 0); }, 0);
}

function formatMeasurementForPrompt_(measurement) {
  if (!measurement) return 'no informada';
  const value = function(item) { return item === '' || item === null || item === undefined ? 'no informada' : item; };
  return 'fecha ' + measurement.date + '; peso ' + value(measurement.weight) + ' kg; grasa abdominal ' + value(measurement.abdominalFat) + '; grasa visceral ' + value(measurement.visceralFat) + '; músculo ' + value(measurement.muscle) + '.';
}

function isRetryableWeeklyGeminiError_(message) {
  return isTransientGeminiError_(message) || /finishReason:\s*MAX_TOKENS/i.test(message || '');
}

function isTransientGeminiError_(message) {
  return /^(HTTP (429|500|502|503|504))\s+con\s+Gemini/i.test(message || '');
}

function buildSafeWeeklyFallback_(patient, measurement, meals, totals) {
  try {
    return buildFallbackRecommendation_(patient, measurement, meals, totals);
  } catch (fallbackError) {
    console.warn('Fallback del reporte semanal falló para ' + patient.username + ': ' + (fallbackError.message || fallbackError));
    const stats = getWeeklyReportStats_(meals);
    return 'Fortalezas: Se registraron ' + meals.length + ' comidas en ' + stats.days + ' días; promedio de ' + stats.averageCalories + ' kcal, ' + stats.averageProtein + ' g de proteínas, ' + stats.averageCarbs + ' g de H.C. y ' + stats.averageFat + ' g de grasas por día con registros.\nAjustes prioritarios: No fue posible generar una recomendación detallada en este momento.\nPróximo paso: Revisá el registro y consultá con tu nutricionista.';
  }
}

function buildFallbackRecommendation_(patient, measurement, meals, totals) {
  const stats = getWeeklyReportStats_(meals);
  const missing = [];
  if (!patient.birthDate) missing.push('fecha de nacimiento');
  if (!patient.height) missing.push('estatura');
  if (!patient.activity) missing.push('nivel de actividad');
  if (!measurement) missing.push('peso y composición corporal');
  const strengths = [];
  const adjustments = [];
  const next = [];
  if (!meals.length) {
    strengths.push('No hay comidas registradas para evaluar la ingesta semanal.');
    adjustments.push('Comenzá cargando cada comida con su descripción, calorías, proteínas, H.C. y grasas estimadas.');
  } else {
    strengths.push('Se registraron ' + meals.length + ' comidas en ' + stats.days + ' días de la semana.');
    if (stats.days >= 6) strengths.push('El registro fue casi diario, lo que permite una comparación más confiable.');
    if (stats.days <= 3) adjustments.push('Faltaron registros en ' + (7 - stats.days) + ' días; completá el historial antes de sacar conclusiones.');
    const minCalories = Number(patient.minCalories);
    const maxCalories = Number(patient.maxCalories);
    const rangeAvailable = patient.minCalories !== '' && patient.maxCalories !== '' && Number.isFinite(minCalories) && Number.isFinite(maxCalories);
    if (rangeAvailable && stats.averageCalories > 0) {
      if (stats.averageCalories < minCalories) adjustments.push('El promedio fue ' + (minCalories - stats.averageCalories) + ' kcal por día por debajo del mínimo informado; revisá primero que los registros estén completos.');
      else if (stats.averageCalories > maxCalories) adjustments.push('El promedio superó en ' + Math.round(stats.averageCalories - maxCalories) + ' kcal por día el máximo informado; priorizá planificación y porciones acordes al objetivo.');
      else strengths.push('El promedio calórico se mantuvo dentro del rango informado en los días con registros.');
    } else adjustments.push('Cargá un rango calórico objetivo para poder comparar el promedio semanal.');
    const macroTargets = [
      ['proteínas', patient.minProtein, patient.maxProtein, stats.averageProtein],
      ['H.C.', patient.minCarbs, patient.maxCarbs, stats.averageCarbs],
      ['grasas', patient.minFat, patient.maxFat, stats.averageFat]
    ];
    macroTargets.forEach(function(item) {
      const label = item[0];
      const minTarget = Number(item[1]);
      const maxTarget = Number(item[2]);
      const average = Number(item[3]);
      const hasMin = item[1] !== '' && Number.isFinite(minTarget);
      const hasMax = item[2] !== '' && Number.isFinite(maxTarget);
      if ((hasMin || hasMax) && average > 0) {
        if (item[1] !== '' && Number.isFinite(minTarget) && average < minTarget) adjustments.push('El promedio de ' + item[0] + ' fue de ' + average + ' g por día con registros, por debajo del mínimo informado de ' + minTarget + ' g; revisá primero que el registro esté completo.');
        else if (item[2] !== '' && Number.isFinite(maxTarget) && average > maxTarget) adjustments.push('El promedio de ' + item[0] + ' fue de ' + average + ' g por día con registros, por encima del máximo informado de ' + maxTarget + ' g; revisá porciones y distribución con tu nutricionista.');
        else if (item[1] !== '' && Number.isFinite(minTarget) && item[2] !== '' && Number.isFinite(maxTarget) && average >= minTarget && average <= maxTarget) strengths.push('El promedio de ' + item[0] + ' se mantuvo dentro del rango informado en los días con registros.');
      }
    });
    adjustments.push('Los promedios registrados fueron ' + stats.averageProtein + ' g de proteínas, ' + stats.averageCarbs + ' g de H.C. y ' + stats.averageFat + ' g de grasas por día; revisá que se distribuyan entre las comidas.');
  }
  next.push(missing.length ? 'Para afinar la próxima evaluación, cargá ' + missing.join(', ') + '.' : 'Mantené el registro diario y revisá los ajustes con tu nutricionista.');
  return formatWeeklyRecommendationSections_({
    fortalezas: strengths.join(' '),
    ajustesPrioritarios: adjustments.join(' '),
    proximoPaso: next.join(' ')
  });
}

function isCompleteRecommendation_(content) {
  const text = String(content || '').trim();
  if (!text || text.length > 900) return false;
  const lines = text.split(/\n+/).map(function(line) { return line.trim(); }).filter(Boolean);
  if (lines.length < 3) return false;
  const hasLabels = /^Fortalezas\s*:/i.test(lines[0]) && /^Ajustes prioritarios\s*:/i.test(lines[1]) && /^Próximo paso\s*:/i.test(lines[2]);
  return hasLabels && lines.slice(0, 3).every(function(line) { return /[.!?]$/.test(line); });
}

function calculateAge_(birthDate, referenceDate) {
  const parts = birthDate.split('-').map(Number);
  let age = referenceDate.getFullYear() - parts[0];
  const beforeBirthday = referenceDate.getMonth() + 1 < parts[1] || (referenceDate.getMonth() + 1 === parts[1] && referenceDate.getDate() < parts[2]);
  return beforeBirthday ? age - 1 : age;
}

function findUser_(sheet, email) {
  return getUsers_().filter(function(user) { return user.username === normalizeUsername_(email); })[0] || null;
}

function getTargetPatient_(user, requestedUsername) {
  if (user.role === 'usuario') return user.username;
  const patientUsername = normalizeUsername_(requestedUsername);
  if (!patientUsername) throw new Error('Seleccioná un paciente.');
  const patient = getUsers_().filter(function(item) {
    return item.username === patientUsername && item.role === 'usuario' && item.nutritionist === user.username;
  })[0];
  if (!patient) throw new Error('Ese paciente no pertenece a tu lista.');
  return patient.username;
}

function requireSession_(token) {
  const raw = token && CacheService.getScriptCache().get('session:' + token);
  if (!raw) throw new Error('La sesión expiró. Volvé a ingresar.');
  return JSON.parse(raw);
}

function requireRole_(token, role) {
  const user = requireSession_(token);
  if (user.role !== role) throw new Error('No tenés permisos para consultar esa información.');
  return user;
}
