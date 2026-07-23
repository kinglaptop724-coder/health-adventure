const SHEET_ID = '1Tj9JQifnNfgHxo9Rn-sxWSwppQwUeUGvMInIAzxu994'; // ชีต 1: ข้อมูลนักเรียน (อ่านอย่างเดียว)
const ACCOUNT_SHEET_ID = '16Kr3NrC0JfafYUaARtp6VINoX9fNdg9gVwk8GBtOvZQ'; // ชีต 2: บัญชีผู้เล่น (อ่านอย่างเดียว)
const SAVE_SHEET_ID = '187uSH9hcYqaHARaHfBpdItcCNY4P4GFNZoO1hVXFzvg'; // ชีต 3: ระบบเซฟเกม (อ่านและเขียน)

function doGet(e) {
  // 📌 ถ้ามีพารามิเตอร์ fn แปลว่านี่คือการเรียก API จากภายนอก (เช่นหน้าเว็บที่ host บน GitHub Pages) ไม่ใช่การเปิดหน้าเกม
  if (e && e.parameter && e.parameter.fn) {
    return handleApiRequest_(e);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('DMC Football Manager')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 📌 ใช้ตอน host หน้าเกมแยกนอก Apps Script (เช่น GitHub Pages) แล้วยิง fetch() มาที่นี่แทน google.script.run
function doPost(e) {
  return handleApiRequest_(e);
}

// รายชื่อฟังก์ชันที่อนุญาตให้เรียกผ่าน API ภายนอกได้ (whitelist ป้องกันไม่ให้เรียกฟังก์ชันอื่นในโปรเจกต์โดยพลการ)
function callWhitelistedFunction_(fn, args) {
  args = args || [];
  switch (fn) {
    case 'getAccountNames': return getAccountNames();
    case 'verifyLogin': return verifyLogin(args[0], args[1]);
    case 'savePlayerData': return savePlayerData(args[0], args[1]);
    case 'getPlayersData': return getPlayersData();
    case 'pingOnline': return pingOnline(args[0]);
    case 'getOnlineStatus': return getOnlineStatus(args[0]);
    case 'sendChallenge': return sendChallenge(args[0], args[1], args[2], args[3]);
    case 'getIncomingChallenge': return getIncomingChallenge(args[0]);
    case 'respondChallenge': return respondChallenge(args[0], args[1], args[2], args[3]);
    case 'getChallengeStatus': return getChallengeStatus(args[0]);
    case 'getMarketState': return getMarketState();
    case 'signPlayerOnMarket': return signPlayerOnMarket(args[0], args[1], args[2]);
    case 'releasePlayerFromMarket': return releasePlayerFromMarket(args[0], args[1]);
    case 'joinLeague': return joinLeague(args[0], args[1], args[2]);
    case 'getLeagueTable': return getLeagueTable(args[0]);
    case 'recordLeagueMatch': return recordLeagueMatch(args[0], args[1], args[2], args[3]);
    case 'adminStartCup': return adminStartCup(args[0]);
    case 'getCupState': return getCupState();
    case 'getMyCupMatch': return getMyCupMatch(args[0]);
    case 'recordCupMatch': return recordCupMatch(args[0], args[1], args[2], args[3]);
    default: throw new Error('ฟังก์ชันนี้ไม่ได้รับอนุญาตให้เรียกผ่าน API ภายนอก: ' + fn);
  }
}

function handleApiRequest_(e) {
  let output;
  try {
    let fn, args;
    if (e.postData && e.postData.contents) {
      // เรียกผ่าน POST (fetch จากหน้าเว็บภายนอก) — เนื้อหาส่งมาเป็น text/plain ที่เป็น JSON เพื่อเลี่ยง CORS preflight
      const body = JSON.parse(e.postData.contents);
      fn = body.fn; args = body.args;
    } else {
      // เรียกผ่าน GET (เผื่อทดสอบผ่านเบราว์เซอร์โดยตรง)
      fn = e.parameter.fn;
      args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
    }
    output = callWhitelistedFunction_(fn, args);
  } catch (err) {
    output = { __gasError: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันดึงรายชื่อ (อ่านจากชีต 2)
function getAccountNames() {
  try {
    const ss = SpreadsheetApp.openById(ACCOUNT_SHEET_ID);
    const sheet = ss.getSheets()[0]; 
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return data.map(row => row[0]).filter(name => name !== ""); 
  } catch (e) {
    return { error: e.toString() };
  }
}

// ฟังก์ชันตรวจสอบล็อกอิน (อ่านชีต 2 -> เขียนชีต 3)
function verifyLogin(playerName, password) {
  try {
    // 1. ดึงข้อมูลจากชีต 2 เพื่อตรวจสอบรหัสผ่าน (อ่านอย่างเดียว)
    const accSS = SpreadsheetApp.openById(ACCOUNT_SHEET_ID);
    const accSheet = accSS.getSheets()[0];
    const accData = accSheet.getDataRange().getValues();
    
    let isValidUser = false;
    
    for (let i = 1; i < accData.length; i++) {
      // คอลัมน์ A คือชื่อ (Index 0), คอลัมน์ D คือ รหัสผ่าน (Index 3)
      if (accData[i][0] === playerName && accData[i][3].toString() === password.toString()) {
        isValidUser = true;
        break;
      }
    }
    
    if (!isValidUser) {
      return { success: false, message: "รหัสผ่านไม่ถูกต้อง หรือไม่พบชื่อผู้ใช้" };
    }

    // 2. ไปจัดการข้อมูลใน ชีต 3 (ชีตเซฟเกม) เพื่อสร้าง Session และดึงข้อมูลเก่า
    const saveSS = SpreadsheetApp.openById(SAVE_SHEET_ID);
    const saveSheet = saveSS.getSheets()[0];
    const saveData = saveSheet.getDataRange().getValues();
    
    let foundRow = -1;
    let savedJson = null;
    
    // ค้นหาว่าผู้เล่นคนนี้เคยมีข้อมูลเซฟในชีต 3 หรือยัง เพื่อป้องกันการสร้างข้อมูลซ้ำ
    for (let i = 0; i < saveData.length; i++) {
      if (saveData[i][0] === playerName) {
        foundRow = i + 1;
        savedJson = saveData[i][3]; // ข้อมูล JSON เซฟเกม เก็บไว้ที่คอลัมน์ D (Index 3)
        break;
      }
    }
    
    const sessionID = Utilities.getUuid();
    const now = new Date();
    
    if (foundRow !== -1) {
      // เคยเล่นแล้ว: อัปเดต Session และ เวลาออนไลน์ล่าสุด ในบรรทัดเดิม (ล็อคข้อมูลไม่ให้ซ้ำซ้อน)
      saveSheet.getRange(foundRow, 2).setValue(sessionID); // คอลัมน์ B: SessionID
      saveSheet.getRange(foundRow, 3).setValue(now);       // คอลัมน์ C: Last Online
    } else {
      // เล่นครั้งแรก: เพิ่มแถวใหม่ลงในชีต 3
      saveSheet.appendRow([playerName, sessionID, now, ""]); 
    }
    
    return { 
      success: true, 
      saveData: savedJson, 
      playerName: playerName,
      message: "เข้าสู่ระบบสำเร็จ" 
    };
    
  } catch (e) {
    return { success: false, message: "ข้อผิดพลาดระบบฐานข้อมูล: " + e.toString() };
  }
}

// ฟังก์ชันบันทึกเกม (เขียนลงชีต 3 เท่านั้น และป้องกันการบันทึกบรรทัดใหม่ซ้ำซ้อน)
function savePlayerData(playerName, jsonData) {
  try {
    const saveSS = SpreadsheetApp.openById(SAVE_SHEET_ID);
    const saveSheet = saveSS.getSheets()[0];
    const data = saveSheet.getDataRange().getValues();
    
    let foundRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === playerName) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow !== -1) {
      // บันทึก JSON เซฟเกม ทับลงที่คอลัมน์ D (Index 3) แถวเดิม
      saveSheet.getRange(foundRow, 4).setValue(jsonData);
    } else {
      // กรณีฉุกเฉิน หากหาไม่เจอจริงๆ ค่อยเพิ่มแถวใหม่
      saveSheet.appendRow([playerName, "", new Date(), jsonData]);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ================== ระบบท้าแข่งสด (Live PvP Challenge) ==================
// ใช้ชีต 3 (SAVE_SHEET_ID) เพิ่มแท็บใหม่ชื่อ "LiveMatches" เก็บคำท้า/สถานะแมตช์
function getLiveMatchSheet_() {
  const ss = SpreadsheetApp.openById(SAVE_SHEET_ID);
  let sh = ss.getSheetByName('LiveMatches');
  if (!sh) {
    sh = ss.insertSheet('LiveMatches');
    sh.appendRow(['matchId', 'challenger', 'opponent', 'status', 'seed', 'challengerOvr', 'opponentOvr', 'startAt', 'createdAt', 'matchType']);
  }
  return sh;
}

// เรียกทุกๆ ~20 วิ จากฝั่ง client เพื่อบอกว่า "ฉันยังออนไลน์อยู่" (heartbeat)
function pingOnline(playerName) {
  try {
    const saveSS = SpreadsheetApp.openById(SAVE_SHEET_ID);
    const saveSheet = saveSS.getSheets()[0];
    const data = saveSheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 0; i < data.length; i++) { if (data[i][0] === playerName) { foundRow = i + 1; break; } }
    const now = new Date();
    if (foundRow !== -1) {
      saveSheet.getRange(foundRow, 5).setValue(now); // คอลัมน์ E: Last Active (heartbeat)
    } else {
      saveSheet.appendRow([playerName, "", now, "", now]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// เช็คว่ารายชื่อเพื่อนที่ส่งมา ใครออนไลน์อยู่บ้างตอนนี้ (heartbeat ภายใน 45 วิที่ผ่านมา)
function getOnlineStatus(names) {
  try {
    const saveSS = SpreadsheetApp.openById(SAVE_SHEET_ID);
    const saveSheet = saveSS.getSheets()[0];
    const data = saveSheet.getDataRange().getValues();
    const now = new Date().getTime();
    const onlineMap = {};
    names.forEach(function(n) { onlineMap[n] = false; });
    for (let i = 0; i < data.length; i++) {
      const name = data[i][0];
      if (onlineMap.hasOwnProperty(name) && data[i][4]) {
        const last = new Date(data[i][4]).getTime();
        if (now - last <= 45000) onlineMap[name] = true;
      }
    }
    return onlineMap;
  } catch (e) {
    return {};
  }
}

// ผู้ท้า: ส่งคำท้าแข่งสดไปหาอีกคน
// matchType: 'league' (ค่าเริ่มต้น ใช้กับท้าแข่งทั่วไป/ลีก) หรือ 'cup' (แมตช์ในถ้วยข้ามลีก — ต้องเป็นคู่ที่จับสลากไว้จริงเท่านั้น)
function sendChallenge(fromPlayer, toPlayer, fromOvr, matchType) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getLiveMatchSheet_();
    const data = sh.getDataRange().getValues();
    // ล้างคำท้าเก่าที่ยังค้าง (pending) ระหว่างคนสองคนนี้ทิ้งก่อน กันชนกัน
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[3] === 'pending' && ((row[1] === fromPlayer && row[2] === toPlayer) || (row[1] === toPlayer && row[2] === fromPlayer))) {
        sh.getRange(i + 1, 4).setValue('cancelled');
      }
    }
    const matchId = Utilities.getUuid();
    const seed = Math.floor(Math.random() * 2147483647);
    const mType = matchType === 'cup' ? 'cup' : 'league';
    sh.appendRow([matchId, fromPlayer, toPlayer, 'pending', seed, fromOvr, '', '', new Date(), mType]);
    return { success: true, matchId: matchId };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ผู้ถูกท้า: เช็คว่ามีคำท้าเข้ามาไหม (โพลทุก ~4 วิ)
function getIncomingChallenge(playerName) {
  try {
    const sh = getLiveMatchSheet_();
    const data = sh.getDataRange().getValues();
    const now = new Date().getTime();
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[2] === playerName && row[3] === 'pending') {
        if (now - new Date(row[8]).getTime() > 25000) { // ค้างเกิน 25 วิ ถือว่าหมดอายุ
          sh.getRange(i + 1, 4).setValue('expired');
          continue;
        }
        return { found: true, matchId: row[0], challenger: row[1], challengerOvr: row[5], matchType: row[9] || 'league' };
      }
    }
    return { found: false };
  } catch (e) {
    return { found: false };
  }
}

// ผู้ถูกท้า: ตอบรับ/ปฏิเสธคำท้า
function respondChallenge(matchId, playerName, accept, myOvr) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getLiveMatchSheet_();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === matchId) {
        if (data[i][3] !== 'pending') return { success: false, message: 'คำท้านี้ไม่สามารถตอบรับได้แล้ว' };
        if (accept) {
          const startAt = new Date().getTime() + 6000; // เผื่อเวลาให้ทั้งสองฝั่งซิงค์กัน 6 วิ
          sh.getRange(i + 1, 4).setValue('accepted');
          sh.getRange(i + 1, 7).setValue(myOvr);
          sh.getRange(i + 1, 8).setValue(startAt);
          return { success: true, startAt: startAt, seed: data[i][4], challengerOvr: data[i][5] };
        } else {
          sh.getRange(i + 1, 4).setValue('declined');
          return { success: true };
        }
      }
    }
    return { success: false, message: 'ไม่พบคำท้านี้' };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ผู้ท้า: เช็คสถานะคำท้าที่ส่งไป (โพลทุก ~2 วิ)
function getChallengeStatus(matchId) {
  try {
    const sh = getLiveMatchSheet_();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === matchId) {
        return { status: data[i][3], seed: data[i][4], opponentOvr: data[i][6], startAt: data[i][7], matchType: data[i][9] || 'league' };
      }
    }
    return { status: 'not_found' };
  } catch (e) {
    return { status: 'error' };
  }
}
// ================== จบระบบท้าแข่งสด ==================

// ================== ตลาดซื้อขายแบบเรียลไทม์ (Global Market Ownership) ==================
// ใช้ชีต 3 (SAVE_SHEET_ID) เพิ่มแท็บใหม่ชื่อ "MarketOwnership" เก็บว่านักเตะแต่ละคน ถูกทีมไหน "เซ็นสัญญา" ไปแล้วบ้าง
// ทำให้นักเตะ 1 คน ถูกเซ็นสัญญาได้แค่ทีมเดียวเท่านั้นในระบบทั้งหมด (กันแย่งกันซื้อซ้ำ)
function getMarketSheet_() {
  const ss = SpreadsheetApp.openById(SAVE_SHEET_ID);
  let sh = ss.getSheetByName('MarketOwnership');
  if (!sh) {
    sh = ss.insertSheet('MarketOwnership');
    sh.appendRow(['studentId', 'ownerName', 'price', 'signedAt']);
  }
  return sh;
}

// ดึงสถานะตลาดทั้งหมดตอนนี้: ใครเซ็นสัญญานักเตะคนไหนไปแล้วบ้าง (โพลทุก ~6 วิ จากฝั่ง client)
// คืนค่าเป็น object { studentId: ownerName, ... } เพื่อให้ payload เล็กและเช็คง่ายฝั่ง client
function getMarketState() {
  try {
    const sh = getMarketSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {};
    const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    const map = {};
    data.forEach(function(row) {
      if (row[0]) map[row[0]] = row[1];
    });
    return map;
  } catch (e) {
    return { error: e.toString() };
  }
}

// เซ็นสัญญานักเตะ (ล็อกกันชนกันตอนสองคนกดพร้อมกัน) — ถ้ามีคนอื่นเซ็นไปก่อนแล้วจะแจ้งกลับทันที
function signPlayerOnMarket(playerName, studentId, price) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getMarketSheet_();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === studentId) {
        if (data[i][1] && data[i][1] !== playerName) {
          return { success: false, message: 'นักเตะคนนี้ถูกทีมอื่นเซ็นสัญญาไปก่อนแล้ว', owner: data[i][1] };
        }
        return { success: true }; // ตัวเองเซ็นไว้อยู่แล้ว (เช่น sync ซ้ำ) ถือว่าสำเร็จ
      }
    }
    sh.appendRow([studentId, playerName, price || '', new Date()]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ปล่อยตัวนักเตะกลับตลาด (ยกเลิกสัญญา) — ลบสิทธิ์ครอบครองออกจากระบบกลาง
function releasePlayerFromMarket(playerName, studentId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getMarketSheet_();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === studentId) {
        if (data[i][1] === playerName) {
          sh.deleteRow(i + 1);
        }
        break;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
// ================== จบระบบตลาดซื้อขายแบบเรียลไทม์ ==================

// ================== ลีกภูมิภาคกลาง (Central Regional League) ==================
// ใช้ชีต 3 (SAVE_SHEET_ID) เพิ่มแท็บใหม่ชื่อ "LeagueCentral" เก็บสมาชิกลีก + สถิติ + กลุ่ม (สาย) ของทุกคนไว้ที่เดียว
// ทำให้ตารางคะแนนที่ผู้เล่นทุกคนในกลุ่มเดียวกันเห็น เป็นตารางเดียวกันจริงๆ (ไม่ใช่คนละจักรวาลเหมือนระบบเดิม)
const LEAGUE_GROUP_SIZE = 22; // จำนวนสมาชิกสูงสุดต่อกลุ่ม/สาย (1 กลุ่ม = 1 ห้องเรียน ~22 คน) ก่อนจะเปิดกลุ่มถัดไปให้อัตโนมัติ

function getLeagueSheet_() {
  const ss = SpreadsheetApp.openById(SAVE_SHEET_ID);
  let sh = ss.getSheetByName('LeagueCentral');
  if (!sh) {
    sh = ss.insertSheet('LeagueCentral');
    sh.appendRow(['playerName', 'groupId', 'ovr', 'played', 'won', 'draw', 'lost', 'gf', 'ga', 'pts', 'updatedAt', 'formation']);
  }
  // 📌 migration: ชีตเก่าที่สร้างไว้ก่อนมีคอลัมน์ 'formation' (สแนปช็อตตำแหน่งผู้เล่นตัวจริง) ให้เติมหัวคอลัมน์ที่ 12 ให้อัตโนมัติ
  if (sh.getLastColumn() < 12) {
    sh.getRange(1, 12).setValue('formation');
  }
  return sh;
}

// เข้าร่วมลีกกลาง (ครั้งแรก) หรืออัปเดตค่าพลังทีม (ovr) ของตัวเอง — เรียกตอนล็อกอิน และทุกครั้งที่เซฟข้อมูล (กันสถิติผู้อื่นใช้ ovr เก่าตอนจำลองแมตช์)
// คนใหม่จะถูกจัดเข้ากลุ่มที่ยังไม่เต็มโดยอัตโนมัติ (กันกลุ่มเดียวใหญ่เกินไปเมื่อนักเรียนมีจำนวนมาก)
// formationJson (ไม่บังคับ): สแนปช็อตทีมตัวจริงปัจจุบันของผู้เล่น เป็นสตริง JSON รูปแบบ
// { "formation": "4-4-2", "players": [ { "name": "...", "type": "GK" }, ... 11 คนตามลำดับตำแหน่งในฟอร์เมชัน ] }
// ใช้แสดงผังทีมจริงของคู่แข่งตอนขึ้นสนามแข่งลีกภูมิภาค ถ้าไม่ส่งมา (undefined) จะไม่ไปทับค่าที่บันทึกไว้เดิม
function joinLeague(playerName, ovr, formationJson) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getLeagueSheet_();
    const data = sh.getDataRange().getValues();
    const now = new Date();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === playerName) {
        sh.getRange(i + 1, 3).setValue(ovr || 0);
        sh.getRange(i + 1, 11).setValue(now);
        if (typeof formationJson !== 'undefined') sh.getRange(i + 1, 12).setValue(formationJson || '');
        return { success: true, groupId: data[i][1] };
      }
    }

    // 📌 คนใหม่จะถูกจัดเข้ากลุ่ม (ห้องเรียน) ที่ยังไม่เต็มโดยอัตโนมัติ กลุ่มละสูงสุด LEAGUE_GROUP_SIZE ทีม
    const groupCounts = {};
    for (let i = 1; i < data.length; i++) {
      const g = data[i][1];
      groupCounts[g] = (groupCounts[g] || 0) + 1;
    }
    let groupId = 1;
    while (groupCounts[groupId] >= LEAGUE_GROUP_SIZE) groupId++;

    sh.appendRow([playerName, groupId, ovr || 0, 0, 0, 0, 0, 0, 0, 0, now, formationJson || '']);
    return { success: true, groupId: groupId };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ดึงตารางคะแนนของกลุ่มตัวเอง — ทุกคนในกลุ่มเดียวกันเรียกฟังก์ชันนี้แล้วได้ผลลัพธ์ตรงกันเป๊ะเสมอ (โพลทุก ~8 วิ)
function getLeagueTable(playerName) {
  try {
    const sh = getLeagueSheet_();
    const data = sh.getDataRange().getValues();
    let myGroup = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === playerName) { myGroup = data[i][1]; break; }
    }
    if (myGroup === null) return { groupId: null, table: [] };

    const table = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === myGroup) {
        // 📌 แปลงสแนปช็อตฟอร์เมชัน (คอลัมน์ 12) กลับเป็น object ให้ฝั่งไคลเอนต์ใช้วาดผังทีมคู่แข่งได้ทันที ไม่ต้องยิงคำขอเพิ่ม
        let formation = null;
        try { formation = data[i][11] ? JSON.parse(data[i][11]) : null; } catch (e) { formation = null; }
        table.push({
          playerName: data[i][0], ovr: data[i][2], played: data[i][3], won: data[i][4],
          draw: data[i][5], lost: data[i][6], gf: data[i][7], ga: data[i][8], pts: data[i][9],
          formation: formation
        });
      }
    }
    return { groupId: myGroup, table: table };
  } catch (e) {
    return { groupId: null, table: [], error: e.toString() };
  }
}

// บันทึกผลแมตช์ลงตารางกลาง แบบล็อกอะตอมมิก อัปเดตสถิติของทั้งสองฝั่งพร้อมกันในการเรียกครั้งเดียว
// ⚠️ สำคัญ: ต้องมีแค่ฝั่งเดียวเท่านั้นที่เรียกฟังก์ชันนี้ต่อ 1 แมตช์ (ฝั่ง "ผู้ริเริ่มแมตช์") กันสถิตินับซ้ำสองเท่า
function recordLeagueMatch(playerA, goalsA, playerB, goalsB) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getLeagueSheet_();
    const data = sh.getDataRange().getValues();
    let rowA = -1, rowB = -1, groupA = null, groupB = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === playerA) { rowA = i + 1; groupA = data[i][1]; }
      if (data[i][0] === playerB) { rowB = i + 1; groupB = data[i][1]; }
    }
    if (rowA === -1 || rowB === -1) return { success: false, message: 'ไม่พบผู้เล่นในลีกกลาง' };
    if (groupA !== groupB) return { success: false, message: 'ผู้เล่นทั้งสองไม่ได้อยู่กลุ่ม/สายเดียวกัน' };

    applyLeagueMatchStats_(sh, rowA, goalsA, goalsB);
    applyLeagueMatchStats_(sh, rowB, goalsB, goalsA);

    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function applyLeagueMatchStats_(sh, row, gf, ga) {
  const vals = sh.getRange(row, 4, 1, 6).getValues()[0]; // played, won, draw, lost, gf, ga
  let played = vals[0] + 1, won = vals[1], draw = vals[2], lost = vals[3];
  const gfTotal = vals[4] + gf, gaTotal = vals[5] + ga;
  if (gf > ga) won++; else if (gf < ga) lost++; else draw++;
  const pts = won * 3 + draw;
  sh.getRange(row, 4, 1, 7).setValues([[played, won, draw, lost, gfTotal, gaTotal, pts]]);
  sh.getRange(row, 11).setValue(new Date());
}
// ================== จบระบบลีกภูมิภาคกลาง ==================

// ================== ฟุตบอลถ้วยข้ามลีก (Cross-League Cup) ==================
// ใช้ชีต 3 (SAVE_SHEET_ID) เพิ่มแท็บใหม่ชื่อ "CupCentral" เก็บผังสาย (bracket) เป็น JSON ก้อนเดียว
// เพราะโครงสร้างสายน็อคเอาท์เปลี่ยนรูปร่างได้ (จำนวนทีมไม่คงที่ตามจำนวนกลุ่ม) การเก็บเป็น JSON ยืดหยุ่นกว่าตารางแถว/คอลัมน์ตายตัว
const CUP_ADMIN_PASSWORD = 'dmc2026cup'; // 🔑 ครู/แอดมินใช้รหัสนี้ตอนกดเริ่มถ้วย — ควรเปลี่ยนเป็นรหัสของตัวเองก่อนใช้งานจริง

function getCupSheet_() {
  const ss = SpreadsheetApp.openById(SAVE_SHEET_ID);
  let sh = ss.getSheetByName('CupCentral');
  if (!sh) {
    sh = ss.insertSheet('CupCentral');
    sh.appendRow(['key', 'value', 'updatedAt']);
  }
  return sh;
}

function loadCupBracket_(sh) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'bracket') {
      try { return JSON.parse(data[i][1]); } catch (e) { return null; }
    }
  }
  return null;
}

function saveCupBracket_(sh, bracket) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'bracket') {
      sh.getRange(i + 1, 2).setValue(JSON.stringify(bracket));
      sh.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }
  sh.appendRow(['bracket', JSON.stringify(bracket), new Date()]);
}

// จัดลำดับ seed แบบมาตรฐาน (seed 1 พบ seed สุดท้าย, seed 2 พบ seed รองสุดท้าย ...)
// เพื่อไม่ให้ทีมอันดับต้นๆ (จากคนละกลุ่ม) มาเจอกันเองตั้งแต่รอบแรก
function seedOrder_(n) {
  let seeds = [1];
  while (seeds.length < n) {
    const size = seeds.length * 2;
    const next = [];
    seeds.forEach(function (s) {
      next.push(s);
      next.push(size + 1 - s);
    });
    seeds = next;
  }
  return seeds;
}

// แอดมิน/ครู กดเริ่มถ้วย: ดึง Top 2 อันดับของทุกกลุ่มในลีกกลาง (เรียงตามแต้ม -> ผลต่างประตู -> ประตูได้) มาจับสายน็อคเอาท์
function adminStartCup(adminPassword) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (adminPassword !== CUP_ADMIN_PASSWORD) {
      return { success: false, message: 'รหัสแอดมินไม่ถูกต้อง' };
    }

    const leagueSh = getLeagueSheet_();
    const leagueData = leagueSh.getDataRange().getValues();
    const groups = {};
    for (let i = 1; i < leagueData.length; i++) {
      const row = leagueData[i];
      const g = row[1];
      if (!groups[g]) groups[g] = [];
      groups[g].push({ playerName: row[0], groupId: g, ovr: row[2], pts: row[9], gf: row[7], ga: row[8] });
    }

    let qualifiers = [];
    Object.keys(groups).forEach(function (g) {
      const table = groups[g].slice().sort(function (a, b) {
        if (b.pts !== a.pts) return b.pts - a.pts;
        const gdA = a.gf - a.ga, gdB = b.gf - b.ga;
        if (gdB !== gdA) return gdB - gdA;
        return b.gf - a.gf;
      });
      qualifiers = qualifiers.concat(table.slice(0, 2)); // Top 2 ของแต่ละกลุ่ม
    });

    if (qualifiers.length < 2) {
      return { success: false, message: 'มีผู้เล่นที่ผ่านเข้ารอบไม่พอสำหรับจัดถ้วย (ต้องมีอย่างน้อย 2 ทีม จากอย่างน้อย 1 กลุ่ม)' };
    }

    // เรียงลำดับตามผลงานรวมข้ามกลุ่ม เพื่อใช้จัด seed ของสาย
    qualifiers.sort(function (a, b) {
      if (b.pts !== a.pts) return b.pts - a.pts;
      return (b.gf - b.ga) - (a.gf - a.ga);
    });

    const n = qualifiers.length;
    let bracketSize = 2;
    while (bracketSize < n) bracketSize *= 2;

    const order = seedOrder_(bracketSize);
    const slots = order.map(function (seedNo) {
      return seedNo <= n ? qualifiers[seedNo - 1] : null; // null = ช่องว่าง (ทีมได้ bye ผ่านฟรี)
    });

    const totalRounds = Math.log(bracketSize) / Math.log(2);
    const rounds = [];

    // รอบแรก: จับคู่ตาม slot ที่ seed ไว้
    const round0 = [];
    for (let i = 0; i < slots.length; i += 2) {
      const p1 = slots[i], p2 = slots[i + 1];
      const match = {
        player1: p1 ? p1.playerName : null, player2: p2 ? p2.playerName : null,
        group1: p1 ? p1.groupId : null, group2: p2 ? p2.groupId : null,
        ovr1: p1 ? p1.ovr : null, ovr2: p2 ? p2.ovr : null,
        score1: null, score2: null, winner: null, status: 'pending'
      };
      if (p1 && !p2) { match.winner = p1.playerName; match.status = 'bye'; }
      if (p2 && !p1) { match.winner = p2.playerName; match.status = 'bye'; }
      round0.push(match);
    }
    rounds.push(round0);

    // รอบถัดๆ ไป: เว้นช่องว่างไว้ก่อน รอผู้ชนะจากรอบก่อนหน้ามาเติม
    for (let r = 1; r < totalRounds; r++) {
      const matchCount = bracketSize / Math.pow(2, r + 1);
      const round = [];
      for (let i = 0; i < matchCount; i++) {
        round.push({ player1: null, player2: null, group1: null, group2: null, ovr1: null, ovr2: null, score1: null, score2: null, winner: null, status: 'pending' });
      }
      rounds.push(round);
    }

    const bracket = { bracketId: Utilities.getUuid(), startedAt: new Date().toISOString(), teams: n, rounds: rounds };
    propagateCupByes_(bracket); // เลื่อนผู้ที่ได้ bye เข้ารอบถัดไปทันที

    const sh = getCupSheet_();
    saveCupBracket_(sh, bracket);
    return { success: true, bracket: bracket };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// เลื่อนผู้ชนะ (รวมถึงทีมที่ได้ bye ผ่านฟรี) เข้าไปเติมช่องว่างของรอบถัดไป ทำวนซ้ำจนไม่มีอะไรเลื่อนได้อีก
function propagateCupByes_(bracket) {
  let changed = true;
  while (changed) {
    changed = false;
    for (let r = 0; r < bracket.rounds.length - 1; r++) {
      const round = bracket.rounds[r];
      const nextRound = bracket.rounds[r + 1];
      for (let i = 0; i < round.length; i++) {
        const m = round[i];
        if (!m.winner) continue;
        const nextMatch = nextRound[Math.floor(i / 2)];
        const slot = (i % 2 === 0) ? 'player1' : 'player2';
        const groupSlot = (i % 2 === 0) ? 'group1' : 'group2';
        const ovrSlot = (i % 2 === 0) ? 'ovr1' : 'ovr2';
        if (!nextMatch[slot]) {
          nextMatch[slot] = m.winner;
          nextMatch[groupSlot] = (m.player1 === m.winner) ? m.group1 : m.group2;
          nextMatch[ovrSlot] = (m.player1 === m.winner) ? m.ovr1 : m.ovr2;
          changed = true;
        }
      }
    }
  }
}

// ดึงผังสายถ้วยปัจจุบันทั้งหมด (ให้ทุกคนเห็นตรงกัน)
function getCupState() {
  try {
    const sh = getCupSheet_();
    const bracket = loadCupBracket_(sh);
    if (!bracket) return { active: false };
    return { active: true, bracket: bracket };
  } catch (e) {
    return { active: false, error: e.toString() };
  }
}

// เช็คว่าผู้เล่นคนนี้ ตอนนี้อยู่จุดไหนของถ้วย: รอคู่แข่งจากรอบก่อนหน้า / พร้อมแข่ง / ตกรอบ / เป็นแชมป์
function getMyCupMatch(playerName) {
  try {
    const sh = getCupSheet_();
    const bracket = loadCupBracket_(sh);
    if (!bracket) return { active: false };
    const totalRounds = bracket.rounds.length;

    for (let r = 0; r < bracket.rounds.length; r++) {
      const round = bracket.rounds[r];
      for (let i = 0; i < round.length; i++) {
        const m = round[i];
        if (m.player1 !== playerName && m.player2 !== playerName) continue;

        if (m.status === 'done' || m.status === 'bye') {
          if (m.winner === playerName) {
            const isFinal = (r === bracket.rounds.length - 1);
            if (isFinal) return { active: true, champion: true, round: r, totalRounds: totalRounds };
            break; // ชนะแล้ว ไปเจอในรอบถัดไป (loop ชั้นนอกจะหาเจอเอง)
          } else {
            const opp = (m.player1 === playerName) ? m.player2 : m.player1;
            return {
              active: true, eliminated: true, round: r, totalRounds: totalRounds, opponent: opp,
              myScore: (m.player1 === playerName ? m.score1 : m.score2),
              oppScore: (m.player1 === playerName ? m.score2 : m.score1)
            };
          }
        }

        const opp = (m.player1 === playerName) ? m.player2 : m.player1;
        if (!opp) return { active: true, waiting: true, round: r, totalRounds: totalRounds };
        const oppOvr = (m.player1 === playerName) ? m.ovr2 : m.ovr1;
        return { active: true, ready: true, round: r, totalRounds: totalRounds, opponent: opp, opponentOvr: oppOvr, opponentGroup: (m.player1 === playerName ? m.group2 : m.group1) };
      }
    }
    return { active: true, eliminated: true, notInBracket: true, totalRounds: totalRounds };
  } catch (e) {
    return { active: false, error: e.toString() };
  }
}

// บันทึกผลแมตช์ถ้วย แล้วเลื่อนผู้ชนะเข้ารอบถัดไปอัตโนมัติ
// ⚠️ เช่นเดียวกับลีก: ต้องมีแค่ฝั่งเดียวเท่านั้นที่เรียกฟังก์ชันนี้ต่อ 1 แมตช์ กันบันทึกซ้ำ
function recordCupMatch(playerName, opponentName, myGoals, oppGoals) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getCupSheet_();
    const bracket = loadCupBracket_(sh);
    if (!bracket) return { success: false, message: 'ยังไม่มีการเริ่มถ้วย' };

    let found = null, roundIdx = -1;
    for (let r = 0; r < bracket.rounds.length && !found; r++) {
      const round = bracket.rounds[r];
      for (let i = 0; i < round.length; i++) {
        const m = round[i];
        if (m.status === 'pending' &&
            ((m.player1 === playerName && m.player2 === opponentName) || (m.player1 === opponentName && m.player2 === playerName))) {
          found = m; roundIdx = r; break;
        }
      }
    }
    if (!found) return { success: false, message: 'ไม่พบแมตช์ที่รอผลนี้ในถ้วย (อาจถูกบันทึกไปแล้ว หรือยังไม่ถึงคิว)' };

    const iAmP1 = found.player1 === playerName;
    found.score1 = iAmP1 ? myGoals : oppGoals;
    found.score2 = iAmP1 ? oppGoals : myGoals;

    let winner;
    if (found.score1 > found.score2) winner = found.player1;
    else if (found.score2 > found.score1) winner = found.player2;
    else {
      // เสมอ: ตัดสินด้วยเลขสุ่มที่ seed คงที่จากคู่แข่งขัน กันสองฝั่งเรียกแล้วได้ผลตัดสินไม่ตรงกัน
      const seed = hashStringToSeed(bracket.bracketId + '|' + roundIdx + '|' + found.player1 + '|' + found.player2);
      winner = (mulberry32(seed)() < 0.5) ? found.player1 : found.player2;
    }
    found.winner = winner;
    found.status = 'done';

    propagateCupByes_(bracket);
    saveCupBracket_(sh, bracket);
    return { success: true, winner: winner };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
// ================== จบระบบฟุตบอลถ้วยข้ามลีก ==================

// 📌 ฟังก์ชันสุ่มแบบกำหนดค่าเริ่มต้นได้ (Seeded Random) เพื่อให้นักเรียนคนเดิมได้ค่าเดิมทุกครั้ง (ตำแหน่ง/สถิติ/ราคาไม่เปลี่ยนตอนรีเฟรช)
function hashStringToSeed(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = (hash << 5) - hash + str.charCodeAt(i);
    hash |= 0; // แปลงเป็น 32-bit integer
  }
  return hash >>> 0; // ทำให้เป็นค่าบวก
}

function mulberry32(seed) {
  return function () {
    seed |= 0;
    seed = (seed + 0x6D2B79F5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// ฟังก์ชันดึงข้อมูลตลาดนักเตะ (อ่านจากชีต 1)
function getPlayersData() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('ข้อมูลนักเรียน DMC');
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 4, lastRow - 1, 7).getDisplayValues();
    const positions = ['GK', 'CB', 'LB', 'RB', 'CDM', 'CM', 'CAM', 'LW', 'RW', 'ST'];
    
    let players = data.map(function(row) {
      // สร้างเลขสุ่มที่คงที่เฉพาะนักเรียนคนนี้ โดยอิงจากรหัสนักเรียน+ชื่อ (คนเดิมได้ค่าเดิมเสมอ)
      const seedKey = row[0] + '|' + row[1] + '|' + row[2] + '|' + row[5] + '|' + row[6];
      const rand = mulberry32(hashStringToSeed(seedKey));

      const randomPos = positions[Math.floor(rand() * positions.length)];
      const speed = Math.floor(rand() * 50) + 50;
      const shooting = Math.floor(rand() * 50) + 50;
      const passing = Math.floor(rand() * 50) + 50;
      const ovr = Math.round((speed + shooting + passing) / 3);
      
      let basePrice;
      if(ovr >= 90) basePrice = Math.floor(rand() * 40) + 80;
      else if (ovr >= 80) basePrice = Math.floor(rand() * 30) + 30;
      else if (ovr >= 70) basePrice = Math.floor(rand() * 15) + 10;
      else if (ovr >= 60) basePrice = Math.floor(rand() * 5) + 3;
      else basePrice = Math.floor(rand() * 2) + 1;
      
      const price = basePrice * 1000000;
      
      return {
        class: row[0], room: row[1], studentId: row[2], gender: row[3], title: row[4],
        fname: row[5], lname: row[6], fullname: row[4] + row[5] + ' ' + row[6],
        position: randomPos, speed: speed, shooting: shooting, passing: passing, ovr: ovr, price: price
      };
    });
    return players.filter(p => p.fname !== "");
  } catch (e) {
    return { error: e.toString() };
  }
}
