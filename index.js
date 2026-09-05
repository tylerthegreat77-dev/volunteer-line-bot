require('dotenv').config();

// บังคับให้ Node.js ใช้ Google DNS ในการค้นหาโดเมน
const dns = require('dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);

const express = require('express');
const line = require('@line/bot-sdk');
const mongoose = require('mongoose');
const path = require('path');
const ExcelJS = require('exceljs');
const axios = require('axios');
const FormData = require('form-data');

// 🔑 ใส่ API Key ของ ImgBB ที่นี่
const IMGBB_API_KEY = process.env.IMGBB_API_KEY || '29a6620a3c2df4d494f53e43d5d94286';

const app = express();

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const client = line.messagingApi 
  ? new line.messagingApi.MessagingApiClient({ channelAccessToken: config.channelAccessToken })
  : new line.Client(config);

// Webhook LINE
app.post('/webhook', line.middleware(config), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then((result) => res.json(result))
    .catch((err) => {
      console.error('Webhook Error:', err);
      res.status(500).end();
    });
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// เชื่อมต่อ MongoDB
mongoose.connect(process.env.MONGODB_URI)
  .then(() => console.log('✅ เชื่อมต่อ MongoDB สำเร็จแล้ว!'))
  .catch((err) => console.error('❌ เชื่อมต่อ MongoDB ผิดพลาด:', err));

// Schema เพิ่มคอลัมน์ imageUrl และ userId สำหรับผูกรูป
const volunteerSchema = new mongoose.Schema({
  userId: String,      // LINE User ID
  studentId: String,
  name: { type: String, default: 'ไม่ระบุชื่อ' },
  hours: Number,
  imageUrl: { type: String, default: '' }, // ลิงก์รูปภาพหลักฐาน
  date: { type: Date, default: Date.now }
});

const Volunteer = mongoose.model('Volunteer', volunteerSchema);

// Schema เก็บรูปชั่วคราวรอคนพิมพ์บันทึก
const tempImageSchema = new mongoose.Schema({
  userId: String,
  imageUrl: String,
  createdAt: { type: Date, default: Date.now, expires: 1800 } // ลบทิ้งอัตโนมัติใน 30 นาที
});

const TempImage = mongoose.model('TempImage', tempImageSchema);

// ==========================================
// 👑 ADMIN DASHBOARD & API
// ==========================================

// 1. API ดึงประวัติรายการจิตอาสาทั้งหมดพร้อมรูปภาพ
app.get('/api/admin/records', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ date: -1 });
    res.json({ success: true, data: records });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 2. API สรุปยอดนักศึกษา
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

// 3. Export Excel พร้อมลิงก์รูปภาพ
app.get('/admin/export-excel', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ date: -1 });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานชั่วโมงจิตอาสา');

    worksheet.columns = [
      { header: 'รหัสนักศึกษา', key: 'studentId', width: 20 },
      { header: 'ชื่อ-นามสกุล', key: 'name', width: 25 },
      { header: 'ชั่วโมงที่บันทึก', key: 'hours', width: 15 },
      { header: 'วันที่บันทึก', key: 'date', width: 22 },
      { header: 'ลิงก์รูปภาพหลักฐาน', key: 'imageUrl', width: 45 }
    ];

    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'DDEBF7' }
    };

    records.forEach(v => {
      worksheet.addRow({
        studentId: v.studentId,
        name: v.name || 'ไม่ระบุชื่อ',
        hours: v.hours,
        date: v.date ? new Date(v.date).toLocaleString('th-TH') : '-',
        imageUrl: v.imageUrl || 'ไม่มีรูปภาพ'
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Volunteer_Hours_Report.xlsx');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error exporting Excel:', error);
    res.status(500).send('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel');
  }
});

// 4. หน้า Admin Dashboard แสดงรูปภาพ
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
        .img-thumb { width: 60px; height: 60px; object-fit: cover; border-radius: 6px; cursor: pointer; border: 1px solid #ddd; }
      </style>
    </head>
    <body>
      <div class="container py-4">
        <div class="d-flex justify-content-between align-items-center mb-4">
          <h2>📊 รายงานชั่วโมงและหลักฐานจิตอาสา</h2>
          <div>
            <a href="/admin/export-excel" class="btn btn-success me-2">📊 ดาวน์โหลด Excel (.xlsx)</a>
            <button class="btn btn-outline-primary" onclick="loadData()">🔄 รีเฟรชข้อมูล</button>
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
                  <th>หลักฐานรูปภาพ</th>
                  <th>รหัสนักศึกษา</th>
                  <th>ชื่อ-นามสกุล</th>
                  <th>ชั่วโมงที่บันทึก</th>
                  <th>วันที่และเวลา</th>
                </tr>
              </thead>
              <tbody id="recordTableBody">
                <tr><td colspan="5" class="text-center">กำลังโหลดข้อมูล...</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Modal สำหรับขยายดูรูปภาพ -->
      <div class="modal fade" id="imageModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered">
          <div class="modal-content">
            <div class="modal-body text-center p-2">
              <img id="modalImg" src="" class="img-fluid rounded" alt="รูปหลักฐาน">
            </div>
          </div>
        </div>
      </div>

      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
      <script>
        let allRecords = [];

        async function loadData() {
          try {
            const res = await fetch('/api/admin/records');
            const result = await res.json();
            if (result.success) {
              allRecords = result.data;
              renderData(allRecords);
            }
          } catch (err) {
            alert('ไม่สามารถดึงข้อมูลได้');
          }
        }

        function renderData(data) {
          const tbody = document.getElementById('recordTableBody');
          tbody.innerHTML = '';

          if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center">ไม่พบข้อมูลบันทึก</td></tr>';
            return;
          }

          data.forEach(item => {
            const dateStr = new Date(item.date).toLocaleString('th-TH');
            const imgHtml = item.imageUrl 
              ? \`<img src="\${item.imageUrl}" class="img-thumb" onclick="showModal('\${item.imageUrl}')">\`
              : \`<span class="badge bg-secondary">ไม่มีรูป</span>\`;

            const tr = document.createElement('tr');
            tr.innerHTML = \`
              <td>\${imgHtml}</td>
              <td><strong>\${item.studentId}</strong></td>
              <td>\${item.name || 'ไม่ระบุชื่อ'}</td>
              <td><span class="badge bg-success fs-6">\${item.hours} ชม.</span></td>
              <td class="text-muted"><small>\${dateStr}</small></td>
            \`;
            tbody.appendChild(tr);
          });
        }

        function showModal(url) {
          document.getElementById('modalImg').src = url;
          new bootstrap.Modal(document.getElementById('imageModal')).show();
        }

        function filterTable() {
          const searchText = document.getElementById('searchInput').value.toLowerCase();
          const filtered = allRecords.filter(item => 
            item.studentId.toLowerCase().includes(searchText) || 
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

// ฟังก์ชันดึงไฟล์รูปจาก LINE แล้วอัปโหลดไป ImgBB
async function handleImageMessage(event) {
  const userId = event.source.userId;
  const messageId = event.message.id;

  try {
    // 1. ดึงไฟล์รูปจาก LINE
    let stream;
    if (client.getMessageContent) {
      stream = await client.getMessageContent(messageId);
    } else {
      stream = await client.getMessageContent(messageId);
    }

    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);

    // 2. ส่งไฟล์ขึ้น ImgBB
    const formData = new FormData();
    formData.append('image', buffer.toString('base64'));

    const response = await axios.post(`https://api.imgbb.com/1/upload?key=${IMGBB_API_KEY}`, formData, {
      headers: formData.getHeaders()
    });

    const imageUrl = response.data.data.url;

    // 3. บันทึกรูปไว้ชั่วคราว รอคนพิมพ์บันทึกชั่วโมง
    await TempImage.findOneAndUpdate(
      { userId },
      { imageUrl, createdAt: new Date() },
      { upsert: true, new: true }
    );

    return replyTextMsg(
      event.replyToken,
      '📷 ได้รับรูปภาพหลักฐานเรียบร้อยแล้วครับ!\n\nกรุณาพิมพ์บันทึกชั่วโมงต่อได้เลย เช่น:\nบันทึก 6501234567 สมชาย ใจดี 4'
    );

  } catch (error) {
    console.error('Error handling image:', error);
    return replyTextMsg(event.replyToken, '❌ ไม่สามารถบันทึกรูปภาพได้ กรุณาลองส่งใหม่อีกครั้ง');
  }
}

async function handleEvent(event) {
  // หากนักศึกษาส่งรูปภาพเข้ามา
  if (event.type === 'message' && event.message.type === 'image') {
    return handleImageMessage(event);
  }

  if (event.type !== 'message' || event.message.type !== 'text') {
    return Promise.resolve(null);
  }

  const userText = event.message.text.trim();
  const userId = event.source.userId;

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
      // ดึงรูปที่เพิ่งอัปโหลดล่าสุด (ถ้ามี)
      const tempImg = await TempImage.findOneAndDelete({ userId });
      const imageUrl = tempImg ? tempImg.imageUrl : '';

      const newRecord = new Volunteer({ userId, studentId, name, hours, imageUrl });
      await newRecord.save();

      const allRecords = await Volunteer.find({ studentId });
      const totalHours = allRecords.reduce((sum, item) => sum + item.hours, 0);

      let replyText = `✅ บันทึกชั่วโมงจิตอาสาสำเร็จ!\n\n` +
                        `👤 ชื่อ: ${name}\n` +
                        `🆔 รหัส: ${studentId}\n` +
                        `⏱ บันทึกเพิ่ม: ${hours} ชั่วโมง\n` +
                        `📊 ชั่วโมงสะสมรวม: ${totalHours} ชั่วโมง`;

      if (imageUrl) {
        replyText += `\n📷 แนบรูปภาพหลักฐานเรียบร้อยแล้ว`;
      } else {
        replyText += `\n⚠️ (บันทึกโดยไม่มีรูปภาพหลักฐาน)`;
      }

      return replyTextMsg(event.replyToken, replyText);

    } catch (error) {
      console.error('Error saving to DB:', error);
      return replyTextMsg(event.replyToken, '❌ เกิดข้อผิดพลาดในการบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง');
    }
  }

  // 3. ข้อความแนะนำการใช้งาน
  const helpText = `👋 ยินดีต้อนรับสู่ระบบบันทึกชั่วโมงจิตอาสา\n\n` +
                   `📌 ขั้นตอนการใช้งาน:\n` +
                   `1️⃣ (ถ้ามี) ส่งรูปภาพหลักฐานการทำกิจกรรม\n` +
                   `2️⃣ พิมพ์บันทึกชั่วโมงตามรูปแบบ:\n` +
                   `บันทึก <รหัส> <ชื่อ-นามสกุล> <ชั่วโมง>\n` +
                   `ตัวอย่าง: บันทึก 6501234567 สมชาย ใจดี 4\n\n` +
                   `🔍 เช็คชั่วโมงสะสม:\n` +
                   `พิมพ์: เช็คชั่วโมง <รหัสนักศึกษา>`;

  return replyTextMsg(event.replyToken, helpText);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server เปิดทำงานแล้วที่ Port ${PORT}`);
});