const odbc = require("odbc");
const axios = require("axios");
const moment = require("moment");

require("dotenv").config();

const dbFilePath = "C:\\Program Files (x86)\\UNIS\\unis.mdb"; 
const webhookURL = process.env.SLACK_WEBHOOK_URL;
const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=${dbFilePath};Uid=${process.env.DB_USERNAME};Pwd=${process.env.DB_PASSWORD};`;

let lastTime = "000000"; // 마지막 출퇴근 감지 시간

async function sendSlackMessage(text, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      await axios.post(webhookURL, { text });
      console.log(`✅ Slack 메시지 전송 완료: ${text}`);
      return;
    } catch (err) {
      console.error(`⚠️ Slack 메시지 전송 실패 (시도 ${i + 1}/${retries}):`, err);
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
  }
}

console.log("📡 출근/퇴근 감시 시스템 시작됨...");
sendSlackMessage("📡 출근/퇴근 감시 시스템 시작됨...");

async function checkDB() {
  const now = new Date();
  const today = moment().format("YYYYMMDD");
  const isMorning = now.getHours() < 12; // 오전이면 출근, 오후면 퇴근
  const attendanceType = isMorning ? "출근" : "퇴근";

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
      let latestTime = lastTime; // 가장 최신 출퇴근 시간을 저장

      for (let row of result) {
        console.log(`✅ ${attendanceType} 감지: ${row.C_Name} ${row.C_Time}`);

        const message = `🚪 [${attendanceType} 알림] ${row.C_Name}님 ${attendanceType}! 시간: ${row.C_Time}`;
        await sendSlackMessage(message);

        // 가장 늦은 시간을 lastTime으로 업데이트
        if (row.C_Time > latestTime) {
          latestTime = row.C_Time;
        }
      }

      lastTime = latestTime; // 마지막 감지된 시간 업데이트
    }

    await connection.close();
  } catch (err) {
    console.error("DB 오류 발생:", err);
  }
}

// ✅ 10초마다 출근/퇴근 감시 실행 (24시간 동작)
setInterval(checkDB, 10000);

// ✅ 매일 00:00에 lastTime 초기화 (새로운 날 시작 시)
setInterval(() => {
  const now = new Date();
  if (now.getHours() === 0 && now.getMinutes() === 0) {
    console.log("🔄 새로운 날이 시작됨, lastTime 초기화");
    lastTime = "000000";
  }
}, 10000);

// **프로세스 종료 감지 후 Slack 알림 전송**
async function handleShutdown() {
  console.log("🛑 시스템 종료 감지됨, Slack 알림 전송 중...");

  try {
    await sendSlackMessage("🔴 출근/퇴근 감시 시스템이 종료됩니다.");
    console.log("📢 Slack 메시지 전송 완료, 3초 대기 후 종료");
  } catch (err) {
    console.error("⚠️ Slack 메시지 전송 실패:", err);
  }

  await new Promise(resolve => setTimeout(resolve, 3000));
  console.log("🚪 안전하게 종료됩니다.");
  process.exit(0);
}

// **Windows & Linux에서 종료 감지**
process.on("SIGINT", handleShutdown);  // Ctrl + C
process.on("SIGTERM", handleShutdown); // PM2 stop/restart/delete

console.log("🟢 출근/퇴근 감시 시스템이 정상적으로 실행 중...");
sendSlackMessage("🟢 출근/퇴근 감시 시스템이 실행되었습니다.");
