const express = require('express');
const line = require('@line/bot-sdk');
const mongoose = require('mongoose');
const dotenv = require('dotenv');
const xlsx = require('xlsx');

dotenv.config();

const app = express();

// Line Bot Config
const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET,
};

const client = new line.Client(lineConfig);

// Connect MongoDB
const MONGO_URI = process.env.MONGO_URI || 'mongodb://localhost:27017/volunteer_db';
mongoose.connect(MONGO_URI)
  .then(() => console.log('MongoDB Connected Successfully'))
  .catch((err) => console.error('MongoDB Connection Error:', err));

// MongoDB Schema & Model
const VolunteerSchema = new mongoose.Schema({
  userId: String,
  studentId: String,
  name: String,
  facultyCode: String,
  facultyName: String,
  activityName: String,
  hours: Number,
  imageUrl: String,
  date: { type: Date, default: Date.now }
});

const Volunteer = mongoose.model('Volunteer', VolunteerSchema);

// Memory state เก็บสถานะกรอกข้อมูลของแต่ละคน
const userStates = {};

const facultyMap = {
  '01': 'คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร',
  '02': 'คณะบริหารธุรกิจและศิลปศาสตร์',
  '03': 'คณะวิศวกรรมศาสตร์'
};

// Webhook for LINE
app.post('/webhook', line.middleware(lineConfig), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then((result) => res.json(result))
    .catch((err) => {
      console.error(err);
      res.status(500).end();
    });
});

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Handle Events
async function handleEvent(event) {
  if (event.type !== 'message') {
    return Promise.resolve(null);
  }

  const userId = event.source.userId;
  const userText = event.message.text ? event.message.text.trim() : '';

  // เช็คคำสั่งเริ่ม
  if (userText === 'บันทึกจิตอาสา') {
    userStates[userId] = { step: 'WAITING_FACULTY' };
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: 'ยินดีต้อนรับสู่ระบบบันทึกจิตอาสาครับ!\nกรุณาเลือกหรือพิมพ์รหัสคณะของคุณ:\n01 = คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร\n02 = คณะบริหารธุรกิจและศิลปศาสตร์\n03 = คณะวิศวกรรมศาสตร์'
    });
  }

  if (userText === 'เช็คชั่วโมง' || userText === 'ตรวจสอบชั่วโมง') {
    const records = await Volunteer.find({ userId });
    if (records.length === 0) {
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: 'ยังไม่พบประวัติการบันทึกจิตอาสาของคุณครับ'
      });
    }

    const totalHours = records.reduce((sum, item) => sum + (item.hours || 0), 0);
    let msg = `📊 ประวัติการบันทึกของคุณ (${records[0].name || ''})\n`;
    msg += `รหัสนักศึกษา: ${records[0].studentId || '-'}\n`;
    msg += `รวมทั้งสิ้น: ${totalHours} ชั่วโมง\n\n`;
    records.forEach((r, idx) => {
      msg += `${idx + 1}. ${r.activityName} (${r.hours} ชม.)\n`;
    });

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: msg
    });
  }

  const state = userStates[userId];

  if (!state) {
    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: 'พิมพ์ "บันทึกจิตอาสา" เพื่อเริ่มบันทึกข้อมูล หรือพิมพ์ "เช็คชั่วโมง" เพื่อดูประวัติย่อครับ'
    });
  }

  // Step 1: รอรับรหัสคณะ
  if (state.step === 'WAITING_FACULTY') {
    if (!facultyMap[userText]) {
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: 'รหัสคณะไม่ถูกต้อง กรุณาพิมพ์เฉพาะเลขรหัสคณะ (01, 02 หรือ 03) ครับ'
      });
    }
    state.facultyCode = userText;
    state.facultyName = facultyMap[userText];
    state.step = 'WAITING_STUDENT_ID';

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `เลือกคณะ: ${state.facultyName}\n\nกรุณากรอก "รหัสนักศึกษา" ของคุณ:`
    });
  }

  // Step 2: รอรับรหัสนักศึกษา
  if (state.step === 'WAITING_STUDENT_ID') {
    state.studentId = userText;
    state.step = 'WAITING_NAME';

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: 'กรุณากรอก "ชื่อ-นามสกุล" ของคุณ:'
    });
  }

  // Step 3: รอรับชื่อ-นามสกุล
  if (state.step === 'WAITING_NAME') {
    state.name = userText;
    state.step = 'WAITING_ACTIVITY';

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: 'กรุณากรอก "ชื่อกิจกรรมจิตอาสา":'
    });
  }

  // Step 4: รอรับชื่อกิจกรรม
  if (state.step === 'WAITING_ACTIVITY') {
    state.activityName = userText;
    state.step = 'WAITING_HOURS';

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: 'กรุณากรอก "จำนวนชั่วโมงจิตอาสา" (เช่น 2 หรือ 3.5):'
    });
  }

  // Step 5: รอรับจำนวนชั่วโมง
  if (state.step === 'WAITING_HOURS') {
    const hours = parseFloat(userText);
    if (isNaN(hours) || hours <= 0) {
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: 'กรุณากรอกจำนวนชั่วโมงเป็นตัวเลขที่ถูกต้อง (เช่น 2 หรือ 3.5):'
      });
    }

    state.hours = hours;
    state.step = 'WAITING_IMAGE';

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: 'กรุณาส่ง "รูปถ่ายหลักฐาน" การทำกิจกรรมจิตอาสาครับ (หรือพิมพ์ "ไม่มี" หากไม่มีรูป):'
    });
  }

  // Step 6: รอรับรูปถ่ายหลักฐาน
  if (state.step === 'WAITING_IMAGE') {
    if (event.message.type === 'image') {
      state.imageUrl = `https://api.line.me/v2/bot/message/${event.message.id}/content`;
    } else if (userText === 'ไม่มี') {
      state.imageUrl = '';
    } else {
      return client.replyMessage(event.replyToken, {
        type: 'text',
        text: 'กรุณาส่งไฟล์รูปภาพ หรือพิมพ์คำว่า "ไม่มี" ครับ'
      });
    }

    // บันทึกลง MongoDB
    const newRecord = new Volunteer({
      userId,
      studentId: state.studentId,
      name: state.name,
      facultyCode: state.facultyCode,
      facultyName: state.facultyName,
      activityName: state.activityName,
      hours: state.hours,
      imageUrl: state.imageUrl
    });

    await newRecord.save();
    delete userStates[userId];

    return client.replyMessage(event.replyToken, {
      type: 'text',
      text: `✅ บันทึกข้อมูลจิตอาสาเรียบร้อยแล้ว!\n\nคณะ: ${newRecord.facultyName}\nนักศึกษา: ${newRecord.name} (${newRecord.studentId})\nกิจกรรม: ${newRecord.activityName}\nจำนวน: ${newRecord.hours} ชั่วโมง`
    });
  }
}

// API Admin List Records
app.get('/api/admin/records', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ date: -1 });
    res.json({ success: true, data: records });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// API Admin Update Record
app.put('/api/admin/records/:id', async (req, res) => {
  try {
    const { facultyCode, studentId, name, activityName, hours } = req.body;
    const facultyName = facultyMap[facultyCode] || 'ไม่ระบุ';

    const updated = await Volunteer.findByIdAndUpdate(req.params.id, {
      facultyCode,
      facultyName,
      studentId,
      name,
      activityName,
      hours: parseFloat(hours)
    }, { new: true });

    res.json({ success: true, data: updated });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// API Admin Delete Record
app.delete('/api/admin/records/:id', async (req, res) => {
  try {
    await Volunteer.findByIdAndDelete(req.params.id);
    res.json({ success: true, message: 'ลบรายการสำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// Export Excel
app.get('/admin/export-excel', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ date: -1 });
    const data = records.map((r, i) => ({
      'ลำดับ': i + 1,
      'วันที่บันทึก': new Date(r.date).toLocaleString('th-TH'),
      'รหัสคณะ': r.facultyCode || '',
      'ชื่อคณะ': r.facultyName || '',
      'รหัสนักศึกษา': r.studentId || '',
      'ชื่อ-นามสกุล': r.name || '',
      'กิจกรรม': r.activityName || '',
      'จำนวนชั่วโมง': r.hours || 0
    }));

    const ws = xlsx.utils.json_to_sheet(data);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, "VolunteerData");

    const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=volunteer_records.xlsx');
    res.send(buf);
  } catch (err) {
    res.status(500).send('Error generating Excel file');
  }
});

// Admin Dashboard Route
app.get('/admin', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Volunteer Admin Dashboard</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
      <script src="https://unpkg.com/lucide@latest"></script>
      <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
      <style>
        :root {
          --bg-body: #f8fafc;
          --card-border-color: #e2e8f0;
          --primary-color: #4f46e5;
          --primary-hover: #4338ca;
        }

        body {
          font-family: 'Sarabun', 'Plus Jakarta Sans', sans-serif;
          background-color: var(--bg-body);
          color: #334155;
        }

        .navbar {
          background: #ffffff;
          border-bottom: 1px solid var(--card-border-color);
        }

        .card {
          border: 1px solid var(--card-border-color);
          border-radius: 16px;
          box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.02);
          background: #ffffff;
        }

        .stat-icon {
          width: 48px;
          height: 48px;
          border-radius: 12px;
          display: flex;
          align-items: center;
          justify-content: center;
        }

        .table > :not(caption) > * > * {
          padding: 1rem 1.25rem;
          border-bottom-color: #f1f5f9;
        }

        .table thead th {
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #64748b;
          font-weight: 600;
          background-color: #f8fafc;
        }

        .img-thumb {
          width: 44px;
          height: 44px;
          object-fit: cover;
          border-radius: 10px;
          cursor: pointer;
          border: 1px solid #e2e8f0;
          transition: transform 0.2s ease;
        }

        .img-thumb:hover {
          transform: scale(1.08);
        }

        .badge-faculty {
          font-size: 0.75rem;
          padding: 0.35em 0.65em;
          border-radius: 6px;
          font-weight: 500;
        }

        .badge-hours {
          background-color: #ecfdf5;
          color: #047857;
          border: 1px solid #a7f3d0;
          font-weight: 600;
          padding: 0.4em 0.8em;
          border-radius: 20px;
        }

        .search-box .form-control, .search-box .form-select {
          border-radius: 10px;
          border: 1px solid #cbd5e1;
          padding: 0.6rem 1rem;
        }

        .btn-custom-primary {
          background-color: var(--primary-color);
          color: white;
          border-radius: 10px;
          padding: 0.6rem 1.2rem;
          font-weight: 500;
        }

        .btn-custom-primary:hover {
          background-color: var(--primary-hover);
          color: white;
        }

        .chart-container {
          position: relative;
          height: 260px;
          width: 100%;
        }
      </style>
    </head>
    <body>

      <!-- Navbar -->
      <nav class="navbar navbar-expand-lg sticky-top py-3">
        <div class="container-fluid px-4">
          <a class="navbar-brand d-flex align-items-center gap-2 fw-bold text-dark" href="#">
            <div class="bg-primary text-white p-2 rounded-3 d-flex align-items-center justify-content-center" style="width:36px; height:36px;">
              <i data-lucide="heart-handshake" style="width:20px;"></i>
            </div>
            <span>Volunteer System Admin</span>
          </a>
          <div class="d-flex gap-2">
            <a href="/admin/export-excel" class="btn btn-outline-success d-flex align-items-center gap-2" style="border-radius:10px;">
              <i data-lucide="file-spreadsheet" style="width:18px;"></i> Export Excel
            </a>
            <button class="btn btn-custom-primary d-flex align-items-center gap-2" onclick="loadData()">
              <i data-lucide="refresh-cw" style="width:18px;"></i> รีเฟรช
            </button>
          </div>
        </div>
      </nav>

      <div class="container-fluid px-4 py-4">
        
        <!-- Stat Cards Summary -->
        <div class="row g-3 mb-4">
          <div class="col-12 col-md-4">
            <div class="card p-3">
              <div class="d-flex align-items-center justify-content-between">
                <div>
                  <div class="text-muted small fw-medium mb-1">จำนวนการบันทึกทั้งหมด</div>
                  <h3 class="fw-bold mb-0" id="statCount">0</h3>
                </div>
                <div class="stat-icon bg-primary-subtle text-primary">
                  <i data-lucide="clipboard-list"></i>
                </div>
              </div>
            </div>
          </div>
          <div class="col-12 col-md-4">
            <div class="card p-3">
              <div class="d-flex align-items-center justify-content-between">
                <div>
                  <div class="text-muted small fw-medium mb-1">ชั่วโมงสะสมรวมทั้งหมด</div>
                  <h3 class="fw-bold mb-0 text-success" id="statHours">0 <small class="fs-6">ชม.</small></h3>
                </div>
                <div class="stat-icon bg-success-subtle text-success">
                  <i data-lucide="clock"></i>
                </div>
              </div>
            </div>
          </div>
          <div class="col-12 col-md-4">
            <div class="card p-3">
              <div class="d-flex align-items-center justify-content-between">
                <div>
                  <div class="text-muted small fw-medium mb-1">นักศึกษาที่เข้าร่วม</div>
                  <h3 class="fw-bold mb-0 text-indigo" id="statStudents">0 <small class="fs-6">คน</small></h3>
                </div>
                <div class="stat-icon bg-warning-subtle text-warning">
                  <i data-lucide="users"></i>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- Analytics Charts -->
        <div class="row g-3 mb-4">
          <div class="col-12 col-lg-5">
            <div class="card p-3 h-100">
              <div class="fw-bold text-dark mb-3 d-flex align-items-center gap-2">
                <i data-lucide="pie-chart" style="width:18px;" class="text-primary"></i> สัดส่วนชั่วโมงสะสมตามคณะ
              </div>
              <div class="chart-container d-flex justify-content-center align-items-center">
                <canvas id="facultyChart"></canvas>
              </div>
            </div>
          </div>
          <div class="col-12 col-lg-7">
            <div class="card p-3 h-100">
              <div class="fw-bold text-dark mb-3 d-flex align-items-center gap-2">
                <i data-lucide="bar-chart-3" style="width:18px;" class="text-primary"></i> Top 5 นักศึกษาที่มีชั่วโมงจิตอาหาสะสมสูงสุด
              </div>
              <div class="chart-container">
                <canvas id="topStudentsChart"></canvas>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Box -->
        <div class="card p-3 mb-4 search-box">
          <div class="row g-3">
            <div class="col-12 col-md-8">
              <div class="input-group">
                <span class="input-group-text bg-white border-end-0 pe-0" style="border-radius: 10px 0 0 10px; border-color: #cbd5e1;">
                  <i data-lucide="search" class="text-muted" style="width:18px;"></i>
                </span>
                <input type="text" id="searchInput" class="form-control border-start-0" style="border-radius: 0 10px 10px 0;" placeholder="ค้นหาด้วย รหัสนักศึกษา, ชื่อ-นามสกุล หรือกิจกรรม..." onkeyup="filterTable()">
              </div>
            </div>
            <div class="col-12 col-md-4">
              <select id="facultyFilter" class="form-select" onchange="filterTable()">
                <option value="">🏛️ แสดงทุกคณะ</option>
                <option value="01">01 - คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร</option>
                <option value="02">02 - คณะบริหารธุรกิจและศิลปศาสตร์</option>
                <option value="03">03 - คณะวิศวกรรมศาสตร์</option>
              </select>
            </div>
          </div>
        </div>

        <!-- Data Table -->
        <div class="card overflow-hidden">
          <div class="table-responsive">
            <table class="table align-middle mb-0">
              <thead>
                <tr>
                  <th>หลักฐาน</th>
                  <th>ข้อมูลนักศึกษา</th>
                  <th>คณะ</th>
                  <th>ชื่อกิจกรรม</th>
                  <th>ชั่วโมง</th>
                  <th>วันที่บันทึก</th>
                  <th class="text-end">จัดการ</th>
                </tr>
              </thead>
              <tbody id="recordTableBody">
                <tr><td colspan="7" class="text-center py-5 text-muted">กำลังโหลดข้อมูล...</td></tr>
              </tbody>
            </table>
          </div>
        </div>

      </div>

      <!-- Modal ขยายรูป -->
      <div class="modal fade" id="imageModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered">
          <div class="modal-content border-0 bg-transparent">
            <div class="modal-body text-center p-0">
              <img id="modalImg" src="" class="img-fluid rounded-4 shadow-lg" alt="รูปหลักฐาน">
            </div>
          </div>
        </div>
      </div>

      <!-- Modal แก้ไขข้อมูล -->
      <div class="modal fade" id="editModal" tabindex="-1">
        <div class="modal-dialog modal-dialog-centered">
          <div class="modal-content border-0 shadow">
            <div class="modal-header">
              <h5 class="modal-title fw-bold">✏️ แก้ไขรายการจิตอาสา</h5>
              <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
              <form id="editForm">
                <input type="hidden" id="editRecordId">
                <div class="mb-3">
                  <label class="form-label small fw-bold text-muted">คณะ</label>
                  <select id="editFacultyCode" class="form-select" required>
                    <option value="01">01 - คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร</option>
                    <option value="02">02 - คณะบริหารธุรกิจและศิลปศาสตร์</option>
                    <option value="03">03 - คณะวิศวกรรมศาสตร์</option>
                  </select>
                </div>
                <div class="mb-3">
                  <label class="form-label small fw-bold text-muted">รหัสนักศึกษา</label>
                  <input type="text" id="editStudentId" class="form-control" required>
                </div>
                <div class="mb-3">
                  <label class="form-label small fw-bold text-muted">ชื่อ-นามสกุล</label>
                  <input type="text" id="editName" class="form-control" required>
                </div>
                <div class="mb-3">
                  <label class="form-label small fw-bold text-muted">ชื่อกิจกรรม</label>
                  <input type="text" id="editActivityName" class="form-control" required>
                </div>
                <div class="mb-3">
                  <label class="form-label small fw-bold text-muted">จำนวนชั่วโมง</label>
                  <input type="number" step="0.5" id="editHours" class="form-control" required min="0.5">
                </div>
              </form>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-light" data-bs-dismiss="modal">ยกเลิก</button>
              <button type="button" class="btn btn-custom-primary" onclick="submitEdit()">บันทึกการเปลี่ยนแปลง</button>
            </div>
          </div>
        </div>
      </div>

      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
      <script>
        let allRecords = [];
        let facultyChartInstance = null;
        let topStudentsChartInstance = null;

        async function loadData() {
          try {
            const res = await fetch('/api/admin/records');
            const result = await res.json();
            if (result.success) {
              allRecords = result.data;
              updateStats(allRecords);
              renderCharts(allRecords);
              renderData(allRecords);
            }
          } catch (err) {
            alert('ไม่สามารถดึงข้อมูลได้');
          }
        }

        function updateStats(data) {
          document.getElementById('statCount').innerText = data.length.toLocaleString();
          
          const totalHours = data.reduce((sum, item) => sum + (item.hours || 0), 0);
          document.getElementById('statHours').innerText = totalHours.toLocaleString();

          const uniqueStudents = new Set(data.map(item => item.studentId)).size;
          document.getElementById('statStudents').innerText = uniqueStudents.toLocaleString();
        }

        function renderCharts(data) {
          const facultyHours = { '01': 0, '02': 0, '03': 0, 'อื่นๆ': 0 };
          data.forEach(item => {
            if (facultyHours[item.facultyCode] !== undefined) {
              facultyHours[item.facultyCode] += item.hours || 0;
            } else {
              facultyHours['อื่นๆ'] += item.hours || 0;
            }
          });

          const facultyCtx = document.getElementById('facultyChart').getContext('2d');
          if (facultyChartInstance) facultyChartInstance.destroy();

          facultyChartInstance = new Chart(facultyCtx, {
            type: 'doughnut',
            data: {
              labels: ['วิทยาศาสตร์ฯ (01)', 'บริหารธุรกิจฯ (02)', 'วิศวกรรมศาสตร์ (03)', 'อื่นๆ'],
              datasets: [{
                data: [facultyHours['01'], facultyHours['02'], facultyHours['03'], facultyHours['อื่นๆ']],
                backgroundColor: ['#10b981', '#6366f1', '#f59e0b', '#94a3b8']
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { position: 'bottom' } }
            }
          });

          const studentMap = {};
          data.forEach(item => {
            const key = item.studentId;
            if (!studentMap[key]) {
              studentMap[key] = { name: item.name || item.studentId, hours: 0 };
            }
            studentMap[key].hours += item.hours || 0;
          });

          const sortedStudents = Object.values(studentMap)
            .sort((a, b) => b.hours - a.hours)
            .slice(0, 5);

          const studentCtx = document.getElementById('topStudentsChart').getContext('2d');
          if (topStudentsChartInstance) topStudentsChartInstance.destroy();

          topStudentsChartInstance = new Chart(studentCtx, {
            type: 'bar',
            data: {
              labels: sortedStudents.map(s => s.name.length > 15 ? s.name.substring(0, 15) + '...' : s.name),
              datasets: [{
                label: 'ชั่วโมงจิตอาหาสะสม',
                data: sortedStudents.map(s => s.hours),
                backgroundColor: '#4f46e5',
                borderRadius: 8
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              scales: { y: { beginAtZero: true } },
              plugins: { legend: { display: false } }
            }
          });
        }

        function renderData(data) {
          const tbody = document.getElementById('recordTableBody');
          tbody.innerHTML = '';

          if (data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="7" class="text-center py-5 text-muted">ไม่พบข้อมูลบันทึก</td></tr>';
            return;
          }

          const facultyColors = {
            '01': 'bg-success-subtle text-success border-success-subtle',
            '02': 'bg-primary-subtle text-primary border-primary-subtle',
            '03': 'bg-warning-subtle text-warning-emphasis border-warning-subtle'
          };

          data.forEach(item => {
            const dateObj = new Date(item.date);
            const dateStr = dateObj.toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: '2-digit' });
            const timeStr = dateObj.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });

            const imgHtml = item.imageUrl 
              ? '<img src="' + item.imageUrl + '" class="img-thumb shadow-sm" onclick="showModal(\'' + item.imageUrl + '\')">'
              : '<span class="badge bg-light text-muted border py-2 px-2" style="font-size:11px;">ไม่มีรูป</span>';

            const facultyClass = facultyColors[item.facultyCode] || 'bg-light text-dark';
            const facultyBadge = item.facultyCode 
              ? '<span class="badge badge-faculty border ' + facultyClass + '">[' + item.facultyCode + '] ' + (item.facultyName || '') + '</span>'
              : '<span class="text-muted small">ไม่ระบุ</span>';

            const tr = document.createElement('tr');
            tr.innerHTML = 
              '<td>' + imgHtml + '</td>' +
              '<td>' +
                '<div class="fw-bold text-dark">' + (item.name || 'ไม่ระบุชื่อ') + '</div>' +
                '<div class="text-muted small">🆔 ' + item.studentId + '</div>' +
              '</td>' +
              '<td>' + facultyBadge + '</td>' +
              '<td><div class="fw-medium text-dark">' + (item.activityName || 'ไม่ระบุกิจกรรม') + '</div></td>' +
              '<td><span class="badge badge-hours">+' + item.hours + ' ชม.</span></td>' +
              '<td>' +
                '<div class="small fw-medium text-dark">' + dateStr + '</div>' +
                '<div class="text-muted" style="font-size: 11px;">' + timeStr + ' น.</div>' +
              '</td>' +
              '<td class="text-end">' +
                '<button class="btn btn-sm btn-light text-primary me-1" onclick=\'openEditModal(' + JSON.stringify(item) + ')\'>' +
                  '<i data-lucide="edit-3" style="width:16px;"></i>' +
                '</button>' +
                '<button class="btn btn-sm btn-light text-danger" onclick="deleteRecord(\'' + item._id + '\')">' +
                  '<i data-lucide="trash-2" style="width:16px;"></i>' +
                '</button>' +
              '</td>';
            tbody.appendChild(tr);
          });

          lucide.createIcons();
        }

        function showModal(url) {
          document.getElementById('modalImg').src = url;
          new bootstrap.Modal(document.getElementById('imageModal')).show();
        }

        function openEditModal(item) {
          document.getElementById('editRecordId').value = item._id;
          document.getElementById('editFacultyCode').value = item.facultyCode || '01';
          document.getElementById('editStudentId').value = item.studentId || '';
          document.getElementById('editName').value = item.name || '';
          document.getElementById('editActivityName').value = item.activityName || '';
          document.getElementById('editHours').value = item.hours || 0;

          new bootstrap.Modal(document.getElementById('editModal')).show();
        }

        async function submitEdit() {
          const id = document.getElementById('editRecordId').value;
          const payload = {
            facultyCode: document.getElementById('editFacultyCode').value,
            studentId: document.getElementById('editStudentId').value,
            name: document.getElementById('editName').value,
            activityName: document.getElementById('editActivityName').value,
            hours: document.getElementById('editHours').value
          };

          try {
            const res = await fetch('/api/admin/records/' + id, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(payload)
            });

            const result = await res.json();
            if (result.success) {
              bootstrap.Modal.getInstance(document.getElementById('editModal')).hide();
              loadData();
            } else {
              alert('เกิดข้อผิดพลาด: ' + result.message);
            }
          } catch (err) {
            alert('ไม่สามารถอัปเดตข้อมูลได้');
          }
        }

        async function deleteRecord(id) {
          if (!confirm('คุณแน่ใจหรือไม่ว่าต้องการลบรายการนี้?')) return;

          try {
            const res = await fetch('/api/admin/records/' + id, { method: 'DELETE' });
            const result = await res.json();
            if (result.success) {
              loadData();
            } else {
              alert('เกิดข้อผิดพลาด: ' + result.message);
            }
          } catch (err) {
            alert('ไม่สามารถลบรายการได้');
          }
        }

        function filterTable() {
          const searchText = document.getElementById('searchInput').value.toLowerCase();
          const selectedFaculty = document.getElementById('facultyFilter').value;

          const filtered = allRecords.filter(item => {
            const matchSearch = item.studentId.toLowerCase().includes(searchText) || 
                                (item.name && item.name.toLowerCase().includes(searchText)) ||
                                (item.facultyName && item.facultyName.toLowerCase().includes(searchText)) ||
                                (item.activityName && item.activityName.toLowerCase().includes(searchText));
            
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});