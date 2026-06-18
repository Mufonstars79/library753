// ═══════════════════════════════════════════════════════════════════
// Библиотека Автоинструктори — Google Apps Script
// Функции:
//   1. syncAll()           — сваля всички файлове от Drive → GitHub
//   2. processEmails()     — проверява Gmail за нови прикачени файлове
//   3. setupTriggers()     — настройва автоматично изпълнение
//
// НАСТРОЙКА (само веднъж):
//   1. Отиди на https://script.google.com → New project
//   2. Постави целия код
//   3. Project Settings → Script Properties → Add:
//        GITHUB_TOKEN  = (твоя GitHub Personal Access Token)
//        GITHUB_REPO   = Mufonstars79/library753
//        EMAIL_LABEL   = библиотека          ← имейл с тази дума в темата
//   4. Изпълни setupTriggers() веднъж
//   5. Изпълни syncAll() веднъж за начален импорт
// ═══════════════════════════════════════════════════════════════════

// ─── Конфигурация ────────────────────────────────────────────────

const SECTIONS = {
  zakoni: {
    driveId: '1amsQXXQNtIM9hkV1zJ-YSmzTN4sTToVo',
    githubDir: 'files/zakoni',
    emailKeyword: 'закони'   // ако темата съдържа тази дума → тук
  },
  metodika: {
    driveId: '1pO2MksCvgWwsRW8Ta1BT1KFBHl-mR2tL',
    githubDir: 'files/metodika',
    emailKeyword: 'методика'
  },
  pedagogika: {
    driveId: '1gY3wd750g1QBRIdlFHVcyWuVpkbN9pt4',
    githubDir: 'files/pedagogika',
    emailKeyword: 'педагогическа'
  },
  etika: {
    driveId: '1FBcKoUyPZsTYj99uaD-zd7s7NPdmu6PF',
    githubDir: 'files/etika',
    emailKeyword: 'етика'
  }
};

// ─── Главна функция: синхронизация Drive → GitHub ────────────────

function syncAll() {
  const props    = PropertiesService.getScriptProperties();
  const token    = props.getProperty('GITHUB_TOKEN');
  const repo     = props.getProperty('GITHUB_REPO') || 'Mufonstars79/library753';

  if (!token) {
    console.error('Липсва GITHUB_TOKEN в Script Properties!');
    return;
  }

  const manifest = {};

  for (const [key, cfg] of Object.entries(SECTIONS)) {
    console.log(`Синхронизирам секция: ${key}`);
    manifest[key] = [];

    try {
      const folder = DriveApp.getFolderById(cfg.driveId);
      const files  = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();

        // Пропускаме папки (показани са само файлове)
        if (file.getMimeType() === MimeType.FOLDER) continue;

        console.log(`  → ${name}`);

        try {
          const blob    = file.getBlob();
          const bytes   = blob.getBytes();
          const b64     = Utilities.base64Encode(bytes);
          const path    = `${cfg.githubDir}/${name}`;

          uploadToGitHub(repo, token, path, b64, `Sync: ${name}`);

          manifest[key].push({
            name:         name,
            path:         path,
            modifiedTime: file.getLastUpdated().toISOString()
          });
        } catch (fileErr) {
          console.warn(`  ⚠ Пропуснат ${name}: ${fileErr.message}`);
        }
      }
    } catch (folderErr) {
      console.error(`Грешка при папка ${key}: ${folderErr.message}`);
    }
  }

  // Качи manifest.json
  const manifestJson = JSON.stringify(manifest, null, 2);
  const manifestB64  = Utilities.base64Encode(manifestJson, Utilities.Charset.UTF_8);
  uploadToGitHub(repo, token, 'files/manifest.json', manifestB64, 'Обнови manifest.json');

  console.log('✅ Синхронизацията завърши!');
}

// ─── Проверка на Gmail за нови прикачени файлове ─────────────────

function processEmails() {
  const props   = PropertiesService.getScriptProperties();
  const token   = props.getProperty('GITHUB_TOKEN');
  const repo    = props.getProperty('GITHUB_REPO') || 'Mufonstars79/library753';
  const keyword = props.getProperty('EMAIL_LABEL')  || 'библиотека';

  if (!token) return;

  // Търси непрочетени имейли с прикачени файлове
  const threads = GmailApp.search(`is:unread has:attachment subject:${keyword}`);
  let changed   = false;

  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      if (!msg.isUnread()) continue;

      const subject     = msg.getSubject().toLowerCase();
      const attachments = msg.getAttachments();
      if (!attachments.length) { msg.markRead(); continue; }

      // Определи в коя секция да отиде по темата на имейла
      let targetKey = 'zakoni'; // по подразбиране
      for (const [key, cfg] of Object.entries(SECTIONS)) {
        if (subject.includes(cfg.emailKeyword.toLowerCase())) {
          targetKey = key;
          break;
        }
      }

      const cfg = SECTIONS[targetKey];

      for (const att of attachments) {
        const name  = att.getName();
        const b64   = Utilities.base64Encode(att.getBytes());
        const path  = `${cfg.githubDir}/${name}`;

        console.log(`Имейл прикачен файл → ${targetKey}/${name}`);

        try {
          // Качи в Drive папката
          const folder = DriveApp.getFolderById(cfg.driveId);
          folder.createFile(att);

          // Качи директно в GitHub
          uploadToGitHub(repo, token, path, b64, `Имейл: ${name}`);
          changed = true;
        } catch(e) {
          console.warn(`  ⚠ Грешка при ${name}: ${e.message}`);
        }
      }

      msg.markRead();
    }
  }

  // Обнови manifest ако са добавени нови файлове
  if (changed) syncAll();
}

// ─── GitHub API помощна функция ───────────────────────────────────

function uploadToGitHub(repo, token, path, base64Content, commitMsg) {
  const apiUrl  = `https://api.github.com/repos/${repo}/contents/${encodeURI(path)}`;
  const headers = {
    'Authorization': `token ${token}`,
    'Accept':        'application/vnd.github.v3+json',
    'Content-Type':  'application/json'
  };

  // Провери дали файлът вече съществува (нужен е SHA за update)
  let sha = null;
  try {
    const check = UrlFetchApp.fetch(apiUrl, { headers, muteHttpExceptions: true });
    if (check.getResponseCode() === 200) {
      sha = JSON.parse(check.getContentText()).sha;
    }
  } catch(_) {}

  const payload = { message: commitMsg, content: base64Content };
  if (sha) payload.sha = sha;

  const response = UrlFetchApp.fetch(apiUrl, {
    method:           'PUT',
    headers,
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`GitHub API ${code}: ${response.getContentText().substring(0, 200)}`);
  }
}

// ─── Настройка на автоматични тригери ────────────────────────────

function setupTriggers() {
  // Изтрий съществуващи тригери за тези функции
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncAll' ||
        t.getHandlerFunction() === 'processEmails') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Синхронизация на всеки 6 часа
  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .everyHours(6)
    .create();

  // Проверка на имейли на всеки 15 минути
  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .everyMinutes(15)
    .create();

  console.log('✅ Тригерите са настроени!');
  console.log('   syncAll      — на всеки 6 часа');
  console.log('   processEmails — на всеки 15 минути');
}
