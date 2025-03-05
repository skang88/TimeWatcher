const odbc = require("odbc");
const axios = require("axios");
const schedule = require("node-schedule");
const moment = require("moment");

require("dotenv").config();

const dbFilePath = "C:\\Program Files (x86)\\UNIS\\unis.mdb"; 
const webhookURL = process.env.SLACK_WEBHOOK_URL;
const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=${dbFilePath};Uid=${process.env.DB_USERNAME};Pwd=${process.env.DB_PASSWORD};`;

let lastTime = '000000';
let intervalJob = null;
let isWatching = false;

async function sendSlackMessage(text, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      await axios.post(webhookURL, { text });
      console.log(`✅ Slack 메시지 전송 완료: ${text}`);
      return;
    } catch (err) {
      console.error(`⚠️ Slack 메시지 전송 실패 (시도 ${i + 1}/${retries}):`, err);
      await new Promise(resolve => setTimeout(resolve, 2000)); // 2초 대기 후 재시도
    }
  }
}

console.log("📡 출근 감시 시스템 대기중...");
sendSlackMessage("📡 출근 감시 시스템 대기중...");

async function checkDB() {
  const today = moment().format('YYYYMMDD');
  const query = `
    SELECT C_Name, C_Date, C_Time 
    FROM tEnter 
    WHERE C_Date = '${today}' 
    AND L_TID = 38 
    AND C_Time > '${lastTime}'
    ORDER BY C_Time ASC
  `;

  try {
    const connection = await odbc.connect(connectionString);
    const result = await connection.query(query);

    if (result.length > 0) {
      for (let row of result) {
        console.log(`✅ 출근 감지: ${row.C_Name} ${row.C_Time}`);
        lastTime = row.C_Time; 

        const message = `🚪 [출근 알림] ${row.C_Name}님 출근! 시간: ${row.C_Time}`;
        await sendSlackMessage(message);
      }
    }
    await connection.close();
  } catch (err) {
    console.error("DB 오류 발생:", err);
  }
}

function startWatcher() {
  if (isWatching) return;
  isWatching = true;

  console.log("🚨 출근 감시 시작 (06:00~09:00)");
  sendSlackMessage("🚨 출근 감시 시작 (06:00~09:00)");
  intervalJob = setInterval(checkDB, 60000);
}

function stopWatcher() {
  if (!isWatching) return;
  isWatching = false;

  console.log("🛑 출근 감시 종료");
  sendSlackMessage("🛑 출근 감시 종료");
  clearInterval(intervalJob);
  intervalJob = null;
}

// 현재 시간이 06:00~09:00 사이인지 확인 후 실행
const now = new Date();
if (now.getHours() >= 6 && now.getHours() < 9) {
  startWatcher();
}

// 06:00 시작, 09:00 종료 스케줄러
schedule.scheduleJob("0 6 * * *", startWatcher);
schedule.scheduleJob("0 9 * * *", stopWatcher);

// 추가: 매 1분마다 시간 체크 후 자동 실행
setInterval(() => {
  const now = new Date();
  if (now.getHours() >= 6 && now.getHours() < 9 && !intervalJob) {
    startWatcher();
  }
}, 60000);

// ✅ **PM2 종료 시 Slack 메시지가 확실히 전송되도록 수정**
async function handleShutdown(signal) {
  console.log(`🛑 ${signal} 감지됨, 종료 중...`);

  try {
    await sendSlackMessage("🔴 *출근 감시 시스템 종료됨!* PM2 프로세스가 종료됩니다.");
    console.log("📢 Slack 메시지 전송 완료, 5초 대기 후 종료");
  } catch (err) {
    console.error("⚠️ Slack 메시지 전송 실패:", err);
  }

  // PM2가 너무 빨리 종료하지 않도록 5초 대기
  await new Promise(resolve => setTimeout(resolve, 5000));
  console.log("🚪 안전하게 종료됩니다.");
  process.exit(0);
}

// PM2에서 프로세스 종료 시 감지
process.on("SIGINT", handleShutdown);  // Ctrl + C
process.on("SIGTERM", handleShutdown); // PM2 stop/restart/delete

// 프로그램 시작 시 Slack 알림 전송
console.log("🟢 출근 감시 시스템이 시작되었습니다...");
sendSlackMessage("🟢 *출근 감시 시스템 시작됨!* PM2에서 앱이 실행되었습니다.");
