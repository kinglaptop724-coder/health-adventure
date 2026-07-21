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
    case 'sendChallenge': return sendChallenge(args[0], args[1], args[2]);
    case 'getIncomingChallenge': return getIncomingChallenge(args[0]);
    case 'respondChallenge': return respondChallenge(args[0], args[1], args[2], args[3]);
    case 'getChallengeStatus': return getChallengeStatus(args[0]);
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
    sh.appendRow(['matchId', 'challenger', 'opponent', 'status', 'seed', 'challengerOvr', 'opponentOvr', 'startAt', 'createdAt']);
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
function sendChallenge(fromPlayer, toPlayer, fromOvr) {
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
    sh.appendRow([matchId, fromPlayer, toPlayer, 'pending', seed, fromOvr, '', '', new Date()]);
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
        return { found: true, matchId: row[0], challenger: row[1], challengerOvr: row[5] };
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
        return { status: data[i][3], seed: data[i][4], opponentOvr: data[i][6], startAt: data[i][7] };
      }
    }
    return { status: 'not_found' };
  } catch (e) {
    return { status: 'error' };
  }
}
// ================== จบระบบท้าแข่งสด ==================

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