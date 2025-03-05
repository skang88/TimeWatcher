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

async function sendSlackMessage(text) {
  try {
    await axios.post(webhookURL, { text });
  } catch (err) {
    console.error("⚠️ 슬랙 메시지 전송 오류:", err);
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

        const message = {
          text: `🚪 [출근 알림] ${row.C_Name}님 출근! 시간: ${row.C_Time}`
        };

        await sendSlackMessage(message.text);
      }
    }
    await connection.close();
  } catch (err) {
    console.error("DB 오류 발생:", err);
  }
}

function startWatcher() {
  console.log("🚨 출근 감시 시작 (06:00~09:00)");
  sendSlackMessage("🚨 출근 감시 시작 (06:00~09:00)");
  intervalJob = setInterval(checkDB, 60000);
}

function stopWatcher() {
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

// 앱이 종료될 때 슬랙 알림 전송
process.on("SIGINT", async () => {
  console.log("⚠️ 출근 감시 시스템이 종료됩니다...");
  try {
    await sendSlackMessage("⚠️ 출근 감시 시스템이 종료되었습니다.");
  } catch (err) {
    console.error("⚠️ 슬랙 메시지 전송 실패:", err);
  }
  process.exit();
});

