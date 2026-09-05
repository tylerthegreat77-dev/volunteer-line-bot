require('dotenv').config();

// บังคับให้ Node.js ใช้ Google DNS
const dns = require('dns');
dns.setServers(['8.8.8.8', '8.8.4.4']);

const express = require('express');
const line = require('@line/bot-sdk');
const mongoose = require('mongoose');
const path = require('path');
const ExcelJS = require('exceljs');
const cloudinary = require('cloudinary').v2;

// 🔑 ตั้งค่า Cloudinary
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME || 'ao9yrwpm',
  api_key: process.env.CLOUDINARY_API_KEY || '999874921286948',
  api_secret: process.env.CLOUDINARY_API_SECRET || 'Kyk5Mk1qlZ2uQ-Vt9QMJOtUr46M'
});

// 🏛️ ตารางแปลงรหัสคณะเป็นชื่อคณะ
const FACULTY_MAP = {
  '01': 'คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร',
  '02': 'คณะบริหารธุรกิจและศิลปศาสตร์',
  '03': 'คณะวิศวกรรมศาสตร์'
};

const app = express();

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

// Messaging API Client (สำหรับส่งข้อความ)
const client = line.messagingApi 
  ? new line.messagingApi.MessagingApiClient({ channelAccessToken: config.channelAccessToken })
  : new line.Client(config);

// Messaging API Blob Client (สำหรับดึงไฟล์รูปภาพใน SDK v8+)
const blobClient = line.messagingApi 
  ? new line.messagingApi.MessagingApiBlobClient({ channelAccessToken: config.channelAccessToken })
  : client;

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

// Schema สำหรับเก็บข้อมูลจิตอาสา
const volunteerSchema = new mongoose.Schema({
  userId: String,      // LINE User ID
  facultyCode: String, // รหัสคณะ (01, 02, 03)
  facultyName: String, // ชื่อคณะ
  studentId: String,
  name: { type: String, default: 'ไม่ระบุชื่อ' },
  hours: Number,
  imageUrl: { type: String, default: '' }, // ลิงก์รูปภาพหลักฐาน
  date: { type: Date, default: Date.now }
});

const Volunteer = mongoose.model('Volunteer', volunteerSchema);

// Schema สำหรับเก็บรูปภาพชั่วคราว รอคำสั่งพิมพ์บันทึก
const tempImageSchema = new mongoose.Schema({
  userId: String,
  imageUrl: String,
  createdAt: { type: Date, default: Date.now, expires: 1800 } // ลบทิ้งอัตโนมัติใน 30 นาที
});

const TempImage = mongoose.model('TempImage', tempImageSchema);

// ==========================================
// 👑 ADMIN DASHBOARD & API
// ==========================================

// 1. API ดึงประวัติรายการจิตอาสาทั้งหมด
app.get('/api/admin/records', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ date: -1 });
    res.json({ success: true, data: records });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 2. Export Excel พร้อมข้อมูลคณะและลิงก์รูปภาพ
app.get('/admin/export-excel', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ date: -1 });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานชั่วโมงจิตอาสา');

    worksheet.columns = [
      { header: 'รหัสคณะ', key: 'facultyCode', width: 12 },
      { header: 'ชื่อคณะ', key: 'facultyName', width: 35 },
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
        facultyCode: v.facultyCode || '-',
        facultyName: v.facultyName || 'ไม่ระบุคณะ',
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

// 3. หน้า Admin Dashboard แสดงข้อมูล
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
        .img-thumb { width: 55px; height: 55px; object-fit: cover; border-radius: 6px; cursor: pointer; border: 1px solid #ddd; }
      </style>
    </head>
    <body>
      <div class="container-fluid px-4 py-4">
        <div class="d-flex justify-content-between align-items-center mb-4">
          <h2>📊 รายงานชั่วโมงและหลักฐานจิตอาสา</h2>
          <div>
            <a href="/admin/export-excel" class="btn btn-success me-2">📊 ดาวน์โหลด Excel (.xlsx)</a>
            <button class="btn btn-outline-primary" onclick="loadData()">🔄 รีเฟรชข้อมูล</button>
          </div>
        </div>

        <div class="card mb-4 p-3">
          <div class="row g-2">
            <div class="col-md-8">
              <input type="text" id="searchInput" class="form-control" placeholder="🔍 ค้นหาด้วย รหัสนักศึกษา, ชื่อ-นามสกุล หรือ คณะ..." onkeyup="filterTable()">
            </div>
            <div class="col-md-4">
              <select id="facultyFilter" class="form-select" onchange="filterTable()">
                <option value="">🏛️ ทุกคณะ</option>
                <option value="01">01 - คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร</option>
                <option value="02">02 - คณะบริหารธุรกิจและศิลปศาสตร์</option>
                <option value="03">03 - คณะวิศวกรรมศาสตร์</option>
              </select>
            </div>
          </div>
        </div>

        <div class="card p-3">
          <div class="table-responsive">
            <table class="table table-hover align-middle">
              <thead class="table-light">
                <tr>
                  <th>หลักฐาน</th>
                  <th>คณะ</th>
                  <th>รหัสนักศึกษา</th>
                  <th>ชื่อ-นามสกุล</th>
                  <th>ชั่วโมงที่บันทึก</th>
                  <th>วันที่และเวลา</th>
                </tr>
              </thead>
              <tbody id="recordTableBody">
                <tr><td colspan="6" class="text-center">กำลังโหลดข้อมูล...</td></tr>
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
            tbody.innerHTML = '<tr><td colspan="6" class="text-center">ไม่พบข้อมูลบันทึก</td></tr>';
            return;
          }

          data.forEach(item => {
            const dateStr = new Date(item.date).toLocaleString('th-TH');
            const imgHtml = item.imageUrl 
              ? \`<img src="\${item.imageUrl}" class="img-thumb" onclick="showModal('\${item.imageUrl}')">\`
              : \`<span class="badge bg-secondary">ไม่มีรูป</span>\`;

            const facultyBadge = item.facultyCode 
              ? \`<span class="badge bg-info text-dark">[\${item.facultyCode}] \${item.facultyName || ''}</span>\`
              : \`<span class="badge bg-light text-muted">ไม่ระบุ</span>\`;

            const tr = document.createElement('tr');
            tr.innerHTML = \`
              <td>\${imgHtml}</td>
              <td>\${facultyBadge}</td>
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
          const selectedFaculty = document.getElementById('facultyFilter').value;

          const filtered = allRecords.filter(item => {
            const matchSearch = item.studentId.toLowerCase().includes(searchText) || 
                                (item.name && item.name.toLowerCase().includes(searchText)) ||
                                (item.facultyName && item.facultyName.toLowerCase().includes(searchText));
            
            const matchFaculty = selectedFaculty === '' || item.facultyCode === selectedFaculty;

            return matchSearch && matchFaculty;
          });

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
  if (client.replyMessage && typeof client.replyMessage === 'function') {
    return client.replyMessage({
      replyToken: replyToken,
      messages: [{ type: 'text', text }]
    });
  }
  return client.replyMessage(replyToken, { type: 'text', text });
}

// ฟังก์ชันดึงไฟล์รูปจาก LINE แล้วอัปโหลดไป Cloudinary
async function handleImageMessage(event) {
  const userId = event.source.userId;
  const messageId = event.message.id;

  try {
    // ดึงไฟล์รูปภาพจาก LINE
    const stream = await blobClient.getMessageContent(messageId);
    
    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    const base64Image = `data:image/jpeg;base64,${buffer.toString('base64')}`;

    // อัปโหลดไฟล์ขึ้น Cloudinary
    const uploadResult = await cloudinary.uploader.upload(base64Image, {
      folder: 'volunteer_proofs'
    });

    const imageUrl = uploadResult.secure_url;

    // บันทึกลิงก์รูปไว้ชั่วคราว
    await TempImage.findOneAndUpdate(
      { userId },
      { imageUrl, createdAt: new Date() },
      { upsert: true, new: true }
    );

    return replyTextMsg(
      event.replyToken,
      '📷 ได้รับรูปภาพหลักฐานเรียบร้อยแล้วครับ!\n\nกรุณาพิมพ์บันทึกชั่วโมงต่อได้เลย เช่น:\nบันทึก 01 6501234567 สมชาย ใจดี 4'
    );

  } catch (error) {
    console.error('Error handling image with Cloudinary:', error);
    return replyTextMsg(event.replyToken, '❌ ไม่สามารถบันทึกรูปภาพได้ กรุณาลองส่งใหม่อีกครั้ง');
  }
}

async function handleEvent(event) {
  // หากเป็นข้อความรูปภาพ
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
      const facultyName = validRecord ? (validRecord.facultyName || 'ไม่ระบุ') : 'ไม่ระบุ';

      const replyText = `📊 สรุปชั่วโมงจิตอาสา\n\n` +
                        `👤 ชื่อ: ${studentName}\n` +
                        `🆔 รหัส: ${studentId}\n` +
                        `🏛 คณะ: ${facultyName}\n` +
                        `📝 บันทึกทั้งหมด: ${records.length} ครั้ง\n` +
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

    if (parts.length < 5) {
      return replyTextMsg(
        event.replyToken, 
        '❌ รูปแบบคำสั่งไม่ถูกต้อง!\nกรุณาพิมพ์: บันทึก <รหัสคณะ> <รหัสนักศึกษา> <ชื่อ-นามสกุล> <จำนวนชั่วโมง>\n\n' +
        '🏛 รหัสคณะ:\n01 = วิทยาศาสตร์และเทคโนโลยีการเกษตร\n02 = บริหารธุรกิจและศิลปศาสตร์\n03 = วิศวกรรมศาสตร์\n\n' +
        'ตัวอย่าง:\nบันทึก 01 6501234567 สมชาย ใจดี 4'
      );
    }

    const facultyCode = parts[1];
    const studentId = parts[2];
    const hours = parseFloat(parts[parts.length - 1]);
    const name = parts.slice(3, parts.length - 1).join(' ').trim();

    if (!FACULTY_MAP[facultyCode]) {
      return replyTextMsg(
        event.replyToken,
        '❌ รหัสคณะไม่ถูกต้อง!\nกรุณาใช้รหัสคณะดังนี้:\n01 = คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร\n02 = คณะบริหารธุรกิจและศิลปศาสตร์\n03 = คณะวิศวกรรมศาสตร์'
      );
    }

    const facultyName = FACULTY_MAP[facultyCode];

    if (!name) {
      return replyTextMsg(event.replyToken, '❌ ไม่พบชื่อ-นามสกุล กรุณาตรวจสอบรูปแบบอีกครั้ง');
    }

    if (isNaN(hours) || hours <= 0) {
      return replyTextMsg(event.replyToken, '❌ จำนวนชั่วโมงต้องเป็นตัวเลขที่มากกว่า 0');
    }

    try {
      const tempImg = await TempImage.findOneAndDelete({ userId });
      const imageUrl = tempImg ? tempImg.imageUrl : '';

      const newRecord = new Volunteer({ 
        userId, 
        facultyCode, 
        facultyName, 
        studentId, 
        name, 
        hours, 
        imageUrl 
      });

      await newRecord.save();

      const allRecords = await Volunteer.find({ studentId });
      const totalHours = allRecords.reduce((sum, item) => sum + item.hours, 0);

      let replyText = `✅ บันทึกชั่วโมงจิตอาสาสำเร็จ!\n\n` +
                        `👤 ชื่อ: ${name}\n` +
                        `🆔 รหัส: ${studentId}\n` +
                        `🏛 คณะ: ${facultyName}\n` +
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
                   `บันทึก <รหัสคณะ> <รหัสประจำตัว> <ชื่อ-นามสกุล> <ชั่วโมง>\n\n` +
                   `🏛 รหัสคณะ:\n` +
                   `01 = คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร\n` +
                   `02 = คณะบริหารธุรกิจและศิลปศาสตร์\n` +
                   `03 = คณะวิศวกรรมศาสตร์\n\n` +
                   `💡 ตัวอย่าง:\nบันทึก 01 6501234567 สมชาย ใจดี 4\n\n` +
                   `🔍 เช็คชั่วโมงสะสม:\n` +
                   `พิมพ์: เช็คชั่วโมง <รหัสนักศึกษา>`;

  return replyTextMsg(event.replyToken, helpText);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server เปิดทำงานแล้วที่ Port ${PORT}`);
});