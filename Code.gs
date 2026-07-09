const USERS = { 'ninda': 'sayangdino', 'dino': 'sayangninda' };

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Dashboard Ninda & Dino').addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ['PocketNinda', 'PocketKita', 'Wishlist', 'Streak', 'UserStats', 'GameState'];
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      let sheet = ss.insertSheet(name);
      if(name === 'PocketNinda') sheet.appendRow(['ID', 'Nama', 'Kategori', 'Jumlah', 'Jenis', 'Tanggal']);
      if(name === 'PocketKita') sheet.appendRow(['ID', 'Nama', 'Jumlah', 'Tanggal']);
      if(name === 'Wishlist') sheet.appendRow(['ID', 'Nama', 'Biaya', 'Lokasi', 'Hari Ke']);
      if(name === 'Streak') sheet.appendRow(['LastNinda', 'LastDino', 'Count', 'LastUpdate']);
      if(name === 'UserStats') {
        sheet.appendRow(['User', 'Wins', 'Board', 'Border', 'Inventory', 'LobbyStatus']);
        sheet.appendRow(['ninda', 0, 'board-0', 'border-0', 'board-0,border-0', 'WAITING']);
        sheet.appendRow(['dino', 0, 'board-0', 'border-0', 'board-0,border-0', 'WAITING']);
      }
      if(name === 'GameState') {
        sheet.appendRow(['Turn', 'TopCard', 'NindaHand', 'DinoHand', 'Status', 'Winner']);
        sheet.appendRow(['', '', '[]', '[]', 'WAITING', '']);
      }
    }
  });
}

function verifyLogin(u, p) {
  if (USERS[u] && USERS[u] === p) { setupSheets(); return { success: true, user: u }; }
  return { success: false, message: 'Username atau Password salah!' };
}

function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  // Mengambil data dari baris ke-2 (index 1), melewati header
  return data.length > 1 ? data.slice(1).filter(r => r[0] !== "") : []; 
}

function saveData(sName, rowData) {
  // rowData harus berisi: [Nama, Kategori, Jumlah, Jenis, Tanggal]
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sName)
    .appendRow([new Date().getTime().toString(), rowData[0], rowData[1], rowData[2], rowData[3], rowData[4]]);
  return true;
}

function handleStrike(user) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Streak');
  if (s.getLastRow() === 1) s.appendRow(['', '', 0, '']);
  const r = s.getRange(2, 1, 1, 4); let [lN, lD, count, lU] = r.getValues()[0];
  lN = (lN instanceof Date) ? lN.toDateString() : lN; lD = (lD instanceof Date) ? lD.toDateString() : lD; lU = (lU instanceof Date) ? lU.toDateString() : lU; count = Number(count) || 0;
  const today = new Date().toDateString(), yesterday = new Date(new Date().setDate(new Date().getDate()-1)).toDateString();
  if (lU && lU !== today && lU !== yesterday) count = 0;
  if (user === 'ninda') lN = today; if (user === 'dino') lD = today;
  if (lN === today && lD === today && lU !== today) { count++; lU = today; }
  r.setValues([[lN, lD, count, lU]]); return { lastNinda: lN, lastDino: lD, count };
}

function getDashboardData() {
  // Pastikan ambil data dengan aman
  const n = getSheetData('PocketNinda') || [];
  const k = getSheetData('PocketKita') || [];
  const w = getSheetData('Wishlist') || [];
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Streak'); 
  
  let str = { count: 0, lastNinda: '', lastDino: '' };
  if (s && s.getLastRow() > 1) {
    let v = s.getRange(2, 1, 1, 4).getValues()[0];
    if (v) {
      str = { 
        lastNinda: (v[0] instanceof Date) ? v[0].toDateString() : (v[0] || ''), 
        lastDino: (v[1] instanceof Date) ? v[1].toDateString() : (v[1] || ''), 
        count: v[2] || 0 
      };
    }
  }

  // Hitung dengan memastikan nilai bukan null/undefined menggunakan Number()
  let nIn = 0, nOut = 0, kTot = 0, wTot = 0;
  n.forEach(r => {
    let val = Number(r[3]) || 0;
    r[2] === 'Pemasukan' ? nIn += val : nOut += val;
  });
  k.forEach(r => kTot += (Number(r[2]) || 0));
  w.forEach(r => wTot += (Number(r[2]) || 0));

  // Mengembalikan object yang lengkap, dijamin tidak null
  return { 
    nindaData: n, 
    kitaData: k, 
    wishlistData: w, 
    totals: { 
      nindaPemasukan: nIn, 
      nindaPengeluaran: nOut, 
      nindaSisa: nIn - nOut, 
      kitaSisa: kTot - wTot, 
      wishlistBiaya: wTot 
    }, 
    streak: str 
  };
}

function getMinigameData(user) {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UserStats'), str = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Streak');
  let ud = st.getDataRange().getValues().find(r => r[0] === user) || [user, 0, 'board-0', 'border-0', '', 'WAITING'];
  return { wins: ud[1], equippedBoard: ud[2], equippedBorder: ud[3], inventory: ud[4].split(','), strikes: (str.getLastRow() > 1) ? str.getRange(2,3).getValue() : 0 };
}

function buyItem(user, type, id, cost) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), str = ss.getSheetByName('Streak'), st = ss.getSheetByName('UserStats');
  let cS = str.getRange(2,3).getValue(); if (cS < cost) return {success: false, message: "Strike tidak cukup!"};
  str.getRange(2,3).setValue(cS - cost);
  let d = st.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === user) {
      let inv = d[i][4] ? d[i][4].split(',') : []; if(!inv.includes(id)) inv.push(id);
      if (type === 'board') st.getRange(i+1, 3).setValue(id); if (type === 'border') st.getRange(i+1, 4).setValue(id);
      st.getRange(i+1, 5).setValue(inv.join(',')); break;
    }
  }
  return {success: true};
}

// ================= UNO ENGINE =================
function randomCard(onlyNumber = false) {
  const colors = ['red', 'yellow', 'green', 'blue'];
  const values = ['1','2','3','4','5','6','7','8','9'];
  if (!onlyNumber) values.push('+2', '+4');
  return { color: colors[Math.floor(Math.random() * colors.length)], value: values[Math.floor(Math.random() * values.length)] };
}

function getLobbyStatus() {
  const d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UserStats').getDataRange().getValues();
  let nR = false, dR = false;
  for (let i=1; i<d.length; i++) { if (d[i][0]==='ninda'&&d[i][5]==='READY') nR = true; if (d[i][0]==='dino'&&d[i][5]==='READY') dR = true; }
  return { nindaReady: nR, dinoReady: dR };
}

function setLobbyReady(user, isReady) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UserStats'); const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === user) { s.getRange(i+1, 6).setValue(isReady ? 'READY' : 'WAITING'); break; }
  
  let status = getLobbyStatus();
  const gs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GameState');
  let currentStatus = gs.getRange(2,5).getValue();

  // FIX BUG KOSONG: Hanya tulis data jika kedua pemain ready DAN game belum dimulai
  if (status.nindaReady && status.dinoReady && currentStatus !== 'PLAYING') { 
    let turn = Math.random() > 0.5 ? 'ninda' : 'dino';
    let top = JSON.stringify(randomCard(true)); // Start tanpa +2/+4
    let nH = JSON.stringify(Array.from({length:5}, () => randomCard())), dH = JSON.stringify(Array.from({length:5}, () => randomCard()));
    gs.getRange(2,1,1,6).setValues([[turn, top, nH, dH, 'PLAYING', '']]);
  }
  return status;
}

// FITUR KELUAR GAME (RESET)
function quitGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const us = ss.getSheetByName('UserStats');
  const ud = us.getDataRange().getValues();
  for(let i=1; i<ud.length; i++) us.getRange(i+1, 6).setValue('WAITING');
  
  const gs = ss.getSheetByName('GameState');
  gs.getRange(2,1,1,6).setValues([['', '', '[]', '[]', 'WAITING', '']]);
  return true;
}

function getGameState(user) {
  const v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GameState').getRange(2,1,1,6).getValues()[0];
  let nH = [], dH = [], topCard = {};
  // Anti Error JSON Parse
  try { nH = JSON.parse(v[2] || '[]'); } catch(e) { nH = []; }
  try { dH = JSON.parse(v[3] || '[]'); } catch(e) { dH = []; }
  try { topCard = JSON.parse(v[1] || '{}'); } catch(e) { topCard = {}; }
  
  return { turn: v[0], topCard: topCard, myHand: (user==='ninda' ? nH : dH), oppCardCount: (user==='ninda' ? dH.length : nH.length), status: v[4], winner: v[5] };
}

function playUnoCard(user, cardIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), gs = ss.getSheetByName('GameState');
  const state = gs.getRange(2,1,1,6).getValues()[0];
  if(state[0] !== user || state[4] !== 'PLAYING') return {success: false, message: 'Bukan giliranmu!'};
  
  let nH = JSON.parse(state[2]), dH = JSON.parse(state[3]), topCard = JSON.parse(state[1]);
  let myHand = user === 'ninda' ? nH : dH, oppHand = user === 'ninda' ? dH : nH;
  let played = myHand[cardIndex];

  if (played.color !== topCard.color && played.value !== topCard.value && played.value !== '+4') return {success: false, message: 'Kartu tidak cocok (Warna atau Angka harus sama)!'};
  
  myHand.splice(cardIndex, 1); 
  
  if (played.value === '+2') for(let i=0; i<2; i++) oppHand.push(randomCard());
  if (played.value === '+4') for(let i=0; i<4; i++) oppHand.push(randomCard());
  
  let newTurn = user === 'ninda' ? 'dino' : 'ninda';
  let newStatus = 'PLAYING', winner = '';

  if (myHand.length === 0) {
    newStatus = 'FINISHED'; winner = user;
    const us = ss.getSheetByName('UserStats'); const ud = us.getDataRange().getValues();
    for (let i=1; i<ud.length; i++) {
      if(ud[i][0] === user) us.getRange(i+1, 2).setValue(Number(ud[i][1])+1); 
      us.getRange(i+1, 6).setValue('WAITING'); 
    }
  }

  gs.getRange(2,1,1,6).setValues([[newTurn, JSON.stringify(played), JSON.stringify(nH), JSON.stringify(dH), newStatus, winner]]);
  return {success: true};
}

function drawUnoCard(user) {
  const gs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GameState');
  const state = gs.getRange(2,1,1,6).getValues()[0];
  if(state[0] !== user || state[4] !== 'PLAYING') return {success: false, message: 'Bukan giliranmu!'};
  
  let nH = JSON.parse(state[2]), dH = JSON.parse(state[3]);
  let myHand = user === 'ninda' ? nH : dH;
  myHand.push(randomCard());
  
  gs.getRange(2,1,1,6).setValues([[user === 'ninda' ? 'dino' : 'ninda', state[1], JSON.stringify(nH), JSON.stringify(dH), 'PLAYING', state[5]]]);
  return {success: true};
}

// ... (Bagian atas Code.gs tetap sama) ...

function playUnoCard(user, cardIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), gs = ss.getSheetByName('GameState');
  const state = gs.getRange(2,1,1,6).getValues()[0];
  if(state[0] !== user || state[4] !== 'PLAYING') return {success: false, message: 'Bukan giliranmu!'};
  
  let nH = JSON.parse(state[2]), dH = JSON.parse(state[3]), topCard = JSON.parse(state[1]);
  let myHand = user === 'ninda' ? nH : dH;
  let played = myHand[cardIndex];

  if (!played || (played.color !== topCard.color && played.value !== topCard.value && played.value !== '+4')) 
    return {success: false, message: 'Kartu tidak cocok!'};
  
  myHand.splice(cardIndex, 1); 
  
  let oppHand = user === 'ninda' ? dH : nH;
  if (played.value === '+2') for(let i=0; i<2; i++) oppHand.push(randomCard());
  if (played.value === '+4') for(let i=0; i<4; i++) oppHand.push(randomCard());
  
  let newTurn = user === 'ninda' ? 'dino' : 'ninda';
  let newStatus = 'PLAYING', winner = '';

  if (myHand.length === 0) {
    newStatus = 'FINISHED'; winner = user;
    const us = ss.getSheetByName('UserStats'); const ud = us.getDataRange().getValues();
    for (let i=1; i<ud.length; i++) {
      if(ud[i][0] === user) us.getRange(i+1, 2).setValue(Number(ud[i][1])+1); 
      // JANGAN RESET LOBBY DI SINI, BIARKAN USER MELIHAT STATUS FINISHED
    }
  }

  gs.getRange(2,1,1,6).setValues([[newTurn, JSON.stringify(played), JSON.stringify(nH), JSON.stringify(dH), newStatus, winner]]);
  return {success: true};
}

// Tambahkan fungsi ini untuk memastikan notifikasi cuma sekali
function markGameAcknowledged() {
  const gs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GameState');
  gs.getRange(2,5).setValue('FINISHED_ACK'); // Tandai bahwa pemenang sudah dibaca
}
