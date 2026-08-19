require('dotenv').config();

// บังคับให้ Node.js ใช้ Google DNS ในการค้นหาโดเมน (ช่วยแก้ ECONNREFUSED)
const dns = require('dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);

const express = require('express');
const line = require('@line/bot-sdk');
const mongoose = require('mongoose');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

// เช็กเวอร์ชันอัตโนมัติรองรับทั้ง SDK v7 และ v8+
const client = line.messagingApi 
  ? new line.messagingApi.MessagingApiClient({ channelAccessToken: config.channelAccessToken })
  : new line.Client(config);

// ==========================================
// 1. LINE BOT WEBHOOK (ต้องวางก่อน express.json)
// ==========================================
app.post('/webhook', line.middleware(config), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then((result) => res.json(result))
    .catch((err) => {
      console.error('Webhook Error:', err);
      res.status(500).end();
    });
});

// Middleware สำหรับ Route ทั่วไป
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// เชื่อมต่อ MongoDB
mongoose.connect(process.env.MONGODB_URI)
  .then(() => console.log('✅ เชื่อมต่อ MongoDB สำเร็จแล้ว!'))
  .catch((err) => console.error('❌ เชื่อมต่อ MongoDB ผิดพลาด:', err));

// Schema สำหรับเก็บข้อมูลจิตอาสา
const volunteerSchema = new mongoose.Schema({
  studentId: String,
  name: { type: String, default: 'ไม่ระบุชื่อ' },
  hours: Number,
  date: { type: Date, default: Date.now }
});

const Volunteer = mongoose.model('Volunteer', volunteerSchema);

// ==========================================
// 👑 ADMIN DASHBOARD & API
// ==========================================

// 1. API ดึงข้อมูลสรุปชั่วโมงของนักศึกษาทุกคน (ดึงชื่อจริงล่าสุดที่ไม่ใช่ "ไม่ระบุชื่อ")
app.get('/api/admin/summary', async (req, res) => {
  try {
    const summary = await Volunteer.aggregate([
      { $sort: { date: -1 } },
      {
        $group: {
          _id: "$studentId",
          name: { $first: "$name" },
          totalHours: { $sum: "$hours" },
          recordCount: { $sum: 1 },
          lastUpdated: { $max: "$date" }
        }
      },
      { $sort: { _id: 1 } }
    ]);
    res.json({ success: true, data: summary });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 2. Route สำหรับ Export ข้อมูลเป็นไฟล์ Excel (.xlsx)
app.get('/admin/export-excel', async (req, res) => {
  try {
    const summary = await Volunteer.aggregate([
      { $sort: { date: -1 } },
      {
        $group: {
          _id: "$studentId",
          name: { $first: "$name" },
          totalHours: { $sum: "$hours" },
          recordCount: { $sum: 1 },
          lastUpdated: { $max: "$date" }
        }
      },
      { $sort: { _id: 1 } }
    ]);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานชั่วโมงจิตอาสา');

    // กำหนดหัวคอลัมน์
    worksheet.columns = [
      { header: 'รหัสนักศึกษา', key: 'studentId', width: 20 },
      { header: 'ชื่อ-นามสกุล', key: 'name', width: 25 },
      { header: 'จำนวนครั้งที่บันทึก', key: 'recordCount', width: 18 },
      { header: 'ชั่วโมงสะสมรวม', key: 'totalHours', width: 18 },
      { header: 'บันทึกล่าสุดเมื่อ', key: 'lastUpdated', width: 22 }
    ];

    // จัดสไตล์หัวตาราง (ตัวหนา + พื้นหลังสีฟ้าอ่อน)
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'DDEBF7' }
    };

    // เพิ่มข้อมูลนักศึกษาลงในไฟล์ Excel
    summary.forEach(item => {
      worksheet.addRow({
        studentId: item._id,
        name: item.name || 'ไม่ระบุชื่อ',
        recordCount: item.recordCount,
        totalHours: item.totalHours,
        lastUpdated: item.lastUpdated ? new Date(item.lastUpdated).toLocaleString('th-TH') : '-'
      });
    });

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename=Volunteer_Hours_Report.xlsx'
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error exporting Excel:', error);
    res.status(500).send('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel');
  }
});

// 3. หน้าเว็บ Dashboard สำหรับอาจารย์ (เข้าผ่าน /admin)
app.get('/admin', (req, res) => {
  res.send(`
    <!DOCTYPE html>
     <html lang="th">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>ระบบจัดการชั่วโมงจิตอาสา - Admin</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600&display=swap" rel="stylesheet">
      <style>
        body { font-family: 'Sarabun', sans-serif; background-color: #f8f9fa; }
        .card { border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        .table-hover tbody tr:hover { background-color: #f1f3f5; }
      </style>
    </head>
    <body>
      <div class="container py-4">
        <div class="d-flex justify-content-between align-items-center mb-4">
          <h2>📊 รายงานชั่วโมงจิตอาสา (สำหรับอาจารย์/เจ้าหน้าที่)</h2>
          <div>
            <a href="/admin/export-excel" class="btn btn-success me-2">📊 ดาวน์โหลด Excel (.xlsx)</a>
            <button class="btn btn-outline-primary" onclick="loadData()">🔄 รีเฟรชข้อมูล</button>
          </div>
        </div>

        <div class="row g-3 mb-4">
          <div class="col-md-4">
            <div class="card bg-primary text-white p-3">
              <h5>จำนวนนักศึกษาทั้งหมด</h5>
              <h3 id="totalStudents">0 คน</h3>
            </div>
          </div>
          <div class="col-md-4">
            <div class="card bg-success text-white p-3">
              <h5>ชั่วโมงจิตอาสารวมทั้งหมด</h5>
              <h3 id="grandTotalHours">0 ชั่วโมง</h3>
            </div>
          </div>
        </div>

        <div class="card mb-4 p-3">
          <input type="text" id="searchInput" class="form-control" placeholder="🔍 ค้นหาด้วย รหัสนักศึกษา หรือ ชื่อ-นามสกุล..." onkeyup="filterTable()">
        </div>

        <div class="card p-3">
          <div class="table-responsive">
            <table class="table table-hover align-middle">
              <thead class="table-light">
                <tr>
                  <th>รหัสนักศึกษา</th>
                  <th>ชื่อ-นามสกุล</th>
                  <th>จำนวนครั้งที่บันทึก</th>
                  <th>ชั่วโมงสะสมรวม</th>
                  <th>บันทึกล่าสุดเมื่อ</th>
                </tr>
              </thead>
              <tbody id="studentTableBody">
                <tr><td colspan="5" class="text-center">กำลังโหลดข้อมูล...</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <script>
        let allData = [];

        async function loadData() {
          try {
            const res = await fetch('/api/admin/summary');
            const result = await res.json();
            if (result.success) {
              allData = result.data;
              renderData(allData);
            }
          } catch (err) {
            alert('ไม่สามารถดึงข้อมูลได้');
          }
        }

        function renderData(data) {
          const tbody = document.getElementById('studentTableBody');
          tbody.innerHTML = '';

          let grandTotal = 0;

          if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center">ไม่พบข้อมูล</td></tr>';
            document.getElementById('totalStudents').innerText = '0 คน';
            document.getElementById('grandTotalHours').innerText = '0 ชั่วโมง';
            return;
          }

          data.forEach(item => {
            grandTotal += item.totalHours;
            const dateStr = new Date(item.lastUpdated).toLocaleString('th-TH');
            
            const tr = document.createElement('tr');
            tr.innerHTML = \`
              <td><strong>\${item._id}</strong></td>
              <td>\${item.name || 'ไม่ระบุชื่อ'}</td>
              <td><span class="badge bg-secondary">\${item.recordCount} ครั้ง</span></td>
              <td><span class="badge bg-success fs-6">\${item.totalHours} ชม.</span></td>
              <td class="text-muted"><small>\${dateStr}</small></td>
            \`;
            tbody.appendChild(tr);
          });

          document.getElementById('totalStudents').innerText = data.length + ' คน';
          document.getElementById('grandTotalHours').innerText = grandTotal + ' ชั่วโมง';
        }

        function filterTable() {
          const searchText = document.getElementById('searchInput').value.toLowerCase();
          const filtered = allData.filter(item => 
            item._id.toLowerCase().includes(searchText) || 
            (item.name && item.name.toLowerCase().includes(searchText))
          );
          renderData(filtered);
        }

        loadData();
      </script>
    </body>
    </html>
  `);
});

// ==========================================
// 💬 LINE BOT HANDLER
// ==========================================

function replyTextMsg(replyToken, text) {
  if (client.replyMessage && client.replyMessage.length === 1) {
    return client.replyMessage({
      replyToken: replyToken,
      messages: [{ type: 'text', text }]
    });
  }
  return client.replyMessage(replyToken, { type: 'text', text });
}

async function handleEvent(event) {
  if (event.type !== 'message' || event.message.type !== 'text') {
    return Promise.resolve(null);
  }

  const userText = event.message.text.trim();

  // 1. คำสั่ง "เช็คชั่วโมง"
  if (userText.startsWith('เช็คชั่วโมง')) {
    const parts = userText.split(/\s+/).filter(p => p.trim() !== '');
    const studentId = parts[1];

    if (!studentId) {
      return replyTextMsg(event.replyToken, '❌ กรุณาระบุรหัสนักศึกษา เช่น:\nเช็คชั่วโมง 6501234567');
    }

    try {
      const records = await Volunteer.find({ studentId: studentId });

      if (records.length === 0) {
        return replyTextMsg(event.replyToken, `🔍 ไม่พบข้อมูลการบันทึกชั่วโมงของรหัส: ${studentId}`);
      }

      const totalHours = records.reduce((sum, item) => sum + item.hours, 0);
      
      // ดึงชื่อจริงที่มีบันทึกไว้ในระบบ
      const validRecord = records.reverse().find(r => r.name && r.name !== 'ไม่ระบุชื่อ');
      const studentName = validRecord ? validRecord.name : (records[0].name || 'ไม่ระบุชื่อ');

      const replyText = `📊 สรุปชั่วโมงจิตอาสา\n\n` +
                        `👤 ชื่อ: ${studentName}\n` +
                        `🆔 รหัส: ${studentId}\n` +
                        `📝 จำนวนครั้งที่บันทึก: ${records.length} ครั้ง\n` +
                        `⏱ ชั่วโมงสะสมรวม: ${totalHours} ชั่วโมง`;

      return replyTextMsg(event.replyToken, replyText);

    } catch (error) {
      console.error('Error fetching data:', error);
      return replyTextMsg(event.replyToken, '❌ เกิดข้อผิดพลาดในการดึงข้อมูล กรุณาลองใหม่อีกครั้ง');
    }
  }

  // 2. คำสั่ง "บันทึก"
  if (userText.startsWith('บันทึก')) {
    const parts = userText.split(/\s+/).filter(p => p.trim() !== '');

    if (parts.length < 4) {
      return replyTextMsg(event.replyToken, '❌ รูปแบบคำสั่งไม่ถูกต้อง!\nกรุณาพิมพ์: บันทึก <รหัสนักศึกษา> <ชื่อ-นามสกุล> <จำนวนชั่วโมง>\n\nตัวอย่าง:\nบันทึก 6501234567 สมชาย ใจดี 4');
    }

    const studentId = parts[1];
    const hours = parseFloat(parts[parts.length - 1]);
    const name = parts.slice(2, parts.length - 1).join(' ').trim();

    if (!name) {
      return replyTextMsg(event.replyToken, '❌ ไม่พบชื่อ-นามสกุล กรุณาตรวจสอบรูปแบบอีกครั้ง');
    }

    if (isNaN(hours) || hours <= 0) {
      return replyTextMsg(event.replyToken, '❌ จำนวนชั่วโมงต้องเป็นตัวเลขที่มากกว่า 0');
    }

    try {
      const newRecord = new Volunteer({ studentId, name, hours });
      await newRecord.save();

      const allRecords = await Volunteer.find({ studentId });
      const totalHours = allRecords.reduce((sum, item) => sum + item.hours, 0);

      const replyText = `✅ บันทึกชั่วโมงจิตอาสาสำเร็จ!\n\n` +
                        `👤 ชื่อ: ${name}\n` +
                        `🆔 รหัส: ${studentId}\n` +
                        `⏱ บันทึกเพิ่ม: ${hours} ชั่วโมง\n` +
                        `📊 ชั่วโมงสะสมรวม: ${totalHours} ชั่วโมง`;

      return replyTextMsg(event.replyToken, replyText);

    } catch (error) {
      console.error('Error saving to DB:', error);
      return replyTextMsg(event.replyToken, '❌ เกิดข้อผิดพลาดในการบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง');
    }
  }

  // 3. ข้อความแนะนำการใช้งาน
  const helpText = `👋 ยินดีต้อนรับสู่ระบบบันทึกชั่วโมงจิตอาสา\n\n` +
                   `📌 คำสั่งที่สามารถใช้งานได้:\n\n` +
                   `1️⃣ บันทึกชั่วโมง:\n` +
                   `พิมพ์: บันทึก <รหัส> <ชื่อ-นามสกุล> <ชั่วโมง>\n` +
                   `ตัวอย่าง: บันทึก 6501234567 สมชาย ใจดี 4\n\n` +
                   `2️⃣ เช็คชั่วโมงสะสม:\n` +
                   `พิมพ์: เช็คชั่วโมง <รหัสนักศึกษา>\n` +
                   `ตัวอย่าง: เช็คชั่วโมง 6501234567`;

  return replyTextMsg(event.replyToken, helpText);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server เปิดทำงานแล้วที่ Port ${PORT}`);
});