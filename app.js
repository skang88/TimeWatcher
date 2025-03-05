const odbc = require("odbc");
const axios = require("axios");
const schedule = require("node-schedule");
const moment = require("moment");

require("dotenv").config();

const dbFilePath = "C:\\Program Files (x86)\\UNIS\\unis.mdb"; // MDB 파일 경로
const webhookURL = process.env.SLACK_WEBHOOK_URL;
const connectionString = `Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=${dbFilePath};Uid=${process.env.DB_USERNAME};Pwd=${process.env.DB_PASSWORD};`;

let lastTime = '000000';
let intervalJob = null;

console.log("📡 출근 감시 시스템 대기중...");

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
        lastTime = row.C_Time; // 마지막 시간 업데이트

        const message = {
          text: `🚪 [출근 알림] ${row.C_Name}님 출근! 시간: ${row.C_Time}`
        };

        await axios.post(webhookURL, message);
      }
    }
    await connection.close();
  } catch (err) {
    console.error("DB 오류 발생:", err);
  }
}

function startWatcher() {
  console.log("🚨 출근 감시 시작 (06:00~09:00)");
  intervalJob = setInterval(checkDB, 1000);
}

function stopWatcher() {
  console.log("🛑 출근 감시 종료");
  clearInterval(intervalJob);
}

console.log("현재 시간: ", new Date().toLocaleString());

const currentTime = new Date();
const targetTime = new Date();
targetTime.setHours(6, 0, 0, 0); // 오늘 7시로 설정

if (currentTime > targetTime) {
  // 이미 지나간 시간이면 바로 실행
  startWatcher();
}

// 매일 06:00에 감시 시작
schedule.scheduleJob("0 6 * * *", startWatcher);

// 매일 09:00에 감시 종료
schedule.scheduleJob("0 9 * * *", stopWatcher);

