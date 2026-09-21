require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const mongoose = require('mongoose');
const line = require('@line/bot-sdk');
const cloudinary = require('cloudinary').v2;
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();

// ==========================================
// 1. CONFIGURATIONS
// ==========================================

// LINE Bot Configuration
const lineConfig = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN || 'YOUR_CHANNEL_ACCESS_TOKEN',
  channelSecret: process.env.LINE_CHANNEL_SECRET || 'YOUR_CHANNEL_SECRET'
};

const client = new line.Client(lineConfig);

// Cloudinary Configuration
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME || 'ao9yrwpm',
  api_key: process.env.CLOUDINARY_API_KEY || '999874921286948',
  api_secret: process.env.CLOUDINARY_API_SECRET || 'Kyk5Mk1qlZ2uQ-Vt9QMJOtUr46M'
});

// MongoDB Connection
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb+srv://admin:1234@cluster0.abcde.mongodb.net/volunteer_db?retryWrites=true&w=majority';
mongoose.connect(MONGODB_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true
})
.then(() => console.log('Connected to MongoDB successfully!'))
.catch(err => console.error('MongoDB connection error:', err));


// ==========================================
// 2. MONGOOSE SCHEMAS & MODELS
// ==========================================

// Main Volunteer Schema
const volunteerSchema = new mongoose.Schema({
  userId: { type: String, required: true },
  displayName: String,
  pictureUrl: String,
  studentId: { type: String, required: true },
  fullName: { type: String, required: true },
  activityName: { type: String, required: true },
  hours: { type: Number, required: true },
  imageUrl: { type: String, required: true },
  status: { type: String, enum: ['pending', 'approved', 'rejected'], default: 'pending' },
  createdAt: { type: Date, default: Date.now }
});

const Volunteer = mongoose.model('Volunteer', volunteerSchema);

// Temp Image Schema for 2-Step Registration (TTL 30 mins)
const tempImageSchema = new mongoose.Schema({
  userId: { type: String, required: true, unique: true },
  imageUrl: { type: String, required: true },
  createdAt: { type: Date, default: Date.now, expires: 1800 }
});

const TempImage = mongoose.model('TempImage', tempImageSchema);


// ==========================================
// 3. MIDDLEWARES & BODY PARSERS
// ==========================================
app.use(express.static('public'));

app.post('/webhook', line.middleware(lineConfig), (req, res) => {
  Promise.all(req.body.events.map(handleEvent))
    .then((result) => res.json(result))
    .catch((err) => {
      console.error('Webhook Error:', err);
      res.status(500).end();
    });
});

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));


// ==========================================
// 4. LINE BOT EVENT HANDLER
// ==========================================
async function handleEvent(event) {
  if (event.type !== 'message') {
    return Promise.resolve(null);
  }

  const userId = event.source.userId;

  // Handle Image Upload
  if (event.message.type === 'image') {
    try {
      const stream = await client.getMessageContent(event.message.id);
      
      const uploadPromise = new Promise((resolve, reject) => {
        const cloudStream = cloudinary.uploader.upload_stream(
          { folder: 'line_volunteer' },
          (error, result) => {
            if (error) return reject(error);
            resolve(result);
          }
        );
        stream.pipe(cloudStream);
      });

      const result = await uploadPromise;
      const imageUrl = result.secure_url;

      await TempImage.findOneAndUpdate(
        { userId: userId },
        { imageUrl: imageUrl, createdAt: new Date() },
        { upsert: true, new: true }
      );

      const replyText = `ได้รับรูปภาพกิจกรรมแล้วครับ! 📸\n\nกรุณาส่งข้อมูลตามรูปแบบนี้เพื่อบันทึกจิตอาสา:\n\n[รหัสนักศึกษา] [ชื่อ-นามสกุล] [ชื่อกิจกรรม] [จำนวนชั่วโมง]\n\nตัวอย่าง:\n65011234 นายสมชาย ใจดี กวาดลานวัด 3`;
      return client.replyMessage(event.replyToken, { type: 'text', text: replyText });

    } catch (err) {
      console.error('Error handling image:', err);
      return client.replyMessage(event.replyToken, { 
        type: 'text', 
        text: 'เกิดข้อผิดพลาดในการอัปโหลดรูปภาพ กรุณาลองใหม่อีกครั้งครับ' 
      });
    }
  }

  // Handle Text Message
  if (event.message.type === 'text') {
    const text = event.message.text.trim();
    
    if (text === 'วิธีใช้งาน' || text === 'ช่วยเหลือ' || text === 'help') {
      const helpText = `📌 วิธีการบันทึกชั่วโมงจิตอาสา:\n1. ส่ง "รูปภาพ" กิจกรรมจิตอาสาเข้ามาในแชท\n2. ส่ง "ข้อความ" ระบุข้อมูลตามรูปแบบนี้:\n   [รหัสนักศึกษา] [ชื่อ-นามสกุล] [ชื่อกิจกรรม] [จำนวนชั่วโมง]\n\nตัวอย่าง:\n65011234 นายสมชาย ใจดี กวาดลานวัด 3`;
      return client.replyMessage(event.replyToken, { type: 'text', text: helpText });
    }

    const tempImg = await TempImage.findOne({ userId: userId });
    if (!tempImg) {
      return client.replyMessage(event.replyToken, { 
        type: 'text', 
        text: '⚠️ ไม่พบรูปภาพกิจกรรมที่ส่งมาก่อนหน้า กรุณาส่ง "รูปภาพกิจกรรม" เข้ามาก่อนส่งข้อความบันทึกครับ' 
      });
    }

    const parts = text.split(/\s+/);
    if (parts.length < 4) {
      return client.replyMessage(event.replyToken, { 
        type: 'text', 
        text: '⚠️ รูปแบบข้อความไม่ถูกต้อง กรุณากรอกให้ครบถ้วน:\n[รหัสนักศึกษา] [ชื่อ-นามสกุล] [ชื่อกิจกรรม] [จำนวนชั่วโมง]\n\nตัวอย่าง:\n65011234 นายสมชาย ใจดี กวาดลานวัด 3' 
      });
    }

    const studentId = parts[0];
    const hoursStr = parts[parts.length - 1];
    const hours = parseFloat(hoursStr);

    if (isNaN(hours) || hours <= 0) {
      return client.replyMessage(event.replyToken, { 
        type: 'text', 
        text: '⚠️ จำนวนชั่วโมงต้องเป็นตัวเลขที่มากกว่า 0 ครับ' 
      });
    }

    const fullName = parts[1];
    const activityName = parts.slice(2, parts.length - 1).join(' ');

    let profile = { displayName: 'LINE User', pictureUrl: '' };
    try {
      profile = await client.getProfile(userId);
    } catch (e) {
      console.log('Could not fetch user profile:', e.message);
    }

    const newRecord = new Volunteer({
      userId: userId,
      displayName: profile.displayName,
      pictureUrl: profile.pictureUrl,
      studentId: studentId,
      fullName: fullName,
      activityName: activityName,
      hours: hours,
      imageUrl: tempImg.imageUrl,
      status: 'pending'
    });

    await newRecord.save();
    await TempImage.deleteOne({ userId: userId });

    const successMessage = `✅ บันทึกข้อมูลจิตอาสาเรียบร้อยแล้ว!\n\n👤 รหัสนักศึกษา: ${studentId}\n🏷️ ชื่อ-นามสกุล: ${fullName}\n📌 กิจกรรม: ${activityName}\n⏳ จำนวน: ${hours} ชั่วโมง\n\nSTATUS: ⏳ รอการอนุมัติจากเจ้าหน้าที่`;
    return client.replyMessage(event.replyToken, { type: 'text', text: successMessage });
  }

  return Promise.resolve(null);
}


// ==========================================
// 5. ADMIN DASHBOARD REST APIs
// ==========================================

// Get All Volunteers List
app.get('/api/admin/volunteers', async (req, res) => {
  try {
    const list = await Volunteer.find().sort({ createdAt: -1 });
    res.json(list);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Update Record Status / Data
app.put('/api/admin/volunteers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const updated = await Volunteer.findByIdAndUpdate(id, req.body, { new: true });
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Delete Record
app.delete('/api/admin/volunteers/:id', async (req, res) => {
  try {
    const { id } = req.params;
    await Volunteer.findByIdAndDelete(id);
    res.json({ message: 'Deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Export Excel API
app.get('/api/admin/export', async (req, res) => {
  try {
    const records = await Volunteer.find().sort({ createdAt: -1 });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Volunteer Records');

    worksheet.columns = [
      { header: 'วันที่บันทึก', key: 'createdAt', width: 20 },
      { header: 'รหัสนักศึกษา', key: 'studentId', width: 15 },
      { header: 'ชื่อ-นามสกุล', key: 'fullName', width: 25 },
      { header: 'กิจกรรม', key: 'activityName', width: 30 },
      { header: 'จำนวนชั่วโมง', key: 'hours', width: 12 },
      { header: 'สถานะ', key: 'status', width: 12 },
      { header: 'ลิงก์รูปภาพ', key: 'imageUrl', width: 45 }
    ];

    records.forEach(r => {
      worksheet.addRow({
        createdAt: new Date(r.createdAt).toLocaleString('th-TH'),
        studentId: r.studentId,
        fullName: r.fullName,
        activityName: r.activityName,
        hours: r.hours,
        status: r.status === 'approved' ? 'อนุมัติ' : r.status === 'rejected' ? 'ปฏิเสธ' : 'รอตรวจสอบ',
        imageUrl: r.imageUrl
      });
    });

    worksheet.getRow(1).font = { bold: true };
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=volunteer_records.xlsx');

    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});


// ==========================================
// 6. ADMIN DASHBOARD FRONTEND ROUTE
// ==========================================
app.get('/admin', (req, res) => {
  res.send(`
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ระบบจัดการจิตอาสา - Admin Dashboard</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <style>
    body { font-family: 'Prompt', sans-serif; background-color: #f4f6f9; color: #333; }
    .navbar { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); }
    .card { border-radius: 12px; border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.05); transition: 0.3s; }
    .stat-card { border-left: 5px solid; }
    .stat-card.primary { border-color: #0d6efd; }
    .stat-card.warning { border-color: #ffc107; }
    .stat-card.success { border-color: #198754; }
    .stat-card.danger { border-color: #dc3545; }
    .table-img { width: 45px; height: 45px; object-fit: cover; border-radius: 8px; cursor: pointer; transition: 0.2s; }
    .table-img:hover { transform: scale(1.1); }
    .student-link { text-decoration: none; font-weight: 600; color: #0d6efd; cursor: pointer; }
    .student-link:hover { text-decoration: underline; color: #0a58ca; }
    .chart-container { position: relative; height: 260px; width: 100%; }
    .avatar-lg { width: 70px; height: 70px; object-fit: cover; border-radius: 50%; border: 3px solid #0d6efd; }
  </style>
</head>
<body>

  <nav class="navbar navbar-dark mb-4">
    <div class="container-fluid px-4">
      <span class="navbar-brand mb-0 h1"><i class="fa-solid fa-hand-holding-heart me-2"></i> ระบบจัดการจิตอาสา Admin Dashboard</span>
      <a href="/api/admin/export" class="btn btn-light btn-sm fw-bold"><i class="fa-solid fa-file-excel text-success me-1"></i> Export Excel</a>
    </div>
  </nav>

  <div class="container-fluid px-4">

    <!-- KPI Cards -->
    <div class="row g-3 mb-4">
      <div class="col-md-3">
        <div class="card stat-card primary p-3">
          <div class="d-flex justify-content-between align-items-center">
            <div>
              <div class="text-muted small">กิจกรรมทั้งหมด</div>
              <h3 class="mb-0 fw-bold" id="kpi-total">0</h3>
            </div>
            <i class="fa-solid fa-list-check fa-2x text-primary opacity-50"></i>
          </div>
        </div>
      </div>
      <div class="col-md-3">
        <div class="card stat-card warning p-3">
          <div class="d-flex justify-content-between align-items-center">
            <div>
              <div class="text-muted small">รอตรวจสอบ</div>
              <h3 class="mb-0 fw-bold text-warning" id="kpi-pending">0</h3>
            </div>
            <i class="fa-solid fa-clock fa-2x text-warning opacity-50"></i>
          </div>
        </div>
      </div>
      <div class="col-md-3">
        <div class="card stat-card success p-3">
          <div class="d-flex justify-content-between align-items-center">
            <div>
              <div class="text-muted small">อนุมัติแล้ว</div>
              <h3 class="mb-0 fw-bold text-success" id="kpi-approved">0</h3>
            </div>
            <i class="fa-solid fa-circle-check fa-2x text-success opacity-50"></i>
          </div>
        </div>
      </div>
      <div class="col-md-3">
        <div class="card stat-card danger p-3">
          <div class="d-flex justify-content-between align-items-center">
            <div>
              <div class="text-muted small">ชั่วโมงอนุมัติรวม</div>
              <h3 class="mb-0 fw-bold text-danger" id="kpi-hours">0</h3>
            </div>
            <i class="fa-solid fa-hourglass-half fa-2x text-danger opacity-50"></i>
          </div>
        </div>
      </div>
    </div>

    <!-- Modern Compact Charts Row -->
    <div class="row g-3 mb-4">
      <div class="col-md-4">
        <div class="card p-3 h-100">
          <h6 class="fw-bold mb-3"><i class="fa-solid fa-chart-pie me-2 text-primary"></i>สัดส่วนสถานะการอนุมัติ</h6>
          <div class="chart-container">
            <canvas id="statusChart"></canvas>
          </div>
        </div>
      </div>
      <div class="col-md-8">
        <div class="card p-3 h-100">
          <h6 class="fw-bold mb-3"><i class="fa-solid fa-chart-column me-2 text-success"></i>สถิติจิตอาสาแยกตามวัน</h6>
          <div class="chart-container">
            <canvas id="dailyChart"></canvas>
          </div>
        </div>
      </div>
    </div>

    <!-- Data Table Card -->
    <div class="card p-3 mb-5">
      <div class="d-flex flex-column flex-md-row justify-content-between align-items-md-center mb-3 gap-2">
        <h5 class="fw-bold mb-0"><i class="fa-solid fa-table me-2"></i>รายการบันทึกจิตอาสา</h5>
        <div class="d-flex gap-2">
          <input type="text" id="searchInput" class="form-control form-control-sm" placeholder="ค้นหา รหัส / ชื่อ / กิจกรรม..." onkeyup="filterData()">
          <select id="statusFilter" class="form-select form-select-sm" style="width: 150px;" onchange="filterData()">
            <option value="all">สถานะทั้งหมด</option>
            <option value="pending">รอตรวจสอบ</option>
            <option value="approved">อนุมัติแล้ว</option>
            <option value="rejected">ปฏิเสธ</option>
          </select>
        </div>
      </div>

      <div class="table-responsive">
        <table class="table table-hover align-middle" id="volunteerTable">
          <thead class="table-light">
            <tr>
              <th>รูปกิจกรรม</th>
              <th>รหัสนักศึกษา</th>
              <th>ชื่อ-นามสกุล</th>
              <th>กิจกรรม</th>
              <th>ชั่วโมง</th>
              <th>วันที่บันทึก</th>
              <th>สถานะ</th>
              <th class="text-center">จัดการ</th>
            </tr>
          </thead>
          <tbody id="tableBody">
            <!-- Data Rows Loaded via JS -->
          </tbody>
        </table>
      </div>
    </div>

  </div>

  <!-- Modal 1: ดูรายละเอียดประวัตินักศึกษารายบุคคล (NEW) -->
  <div class="modal fade" id="studentDetailModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-lg modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header bg-primary text-white">
          <h5 class="modal-title fw-bold"><i class="fa-solid fa-id-card me-2"></i>ข้อมูลประวัติติตอาสารายบุคคล</h5>
          <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
        </div>
        <div class="modal-body p-4">
          <div class="d-flex align-items-center gap-3 mb-4 p-3 bg-light rounded-3">
            <img id="detailAvatar" src="https://via.placeholder.com/70" class="avatar-lg" alt="Profile">
            <div>
              <h4 class="mb-1 fw-bold" id="detailFullName">-</h4>
              <p class="mb-0 text-muted">รหัสนักศึกษา: <span id="detailStudentId" class="fw-semibold text-dark">-</span></p>
            </div>
            <div class="ms-auto text-end bg-white p-3 rounded-3 shadow-sm border">
              <span class="text-muted small d-block">ชั่วโมงอนุมัติรวม</span>
              <span class="h3 fw-bold text-success mb-0" id="detailTotalHours">0</span>
              <span class="small text-muted"> ชม.</span>
            </div>
          </div>

          <h6 class="fw-bold mb-3"><i class="fa-solid fa-history me-2 text-primary"></i>ประวัติการบันทึกกิจกรรมทั้งหมด</h6>
          <div class="table-responsive">
            <table class="table table-bordered align-middle">
              <thead class="table-light">
                <tr>
                  <th style="width: 70px;">รูปภาพ</th>
                  <th>ชื่อกิจกรรม</th>
                  <th class="text-center" style="width: 90px;">ชั่วโมง</th>
                  <th style="width: 140px;">วันที่บันทึก</th>
                  <th class="text-center" style="width: 110px;">สถานะ</th>
                </tr>
              </thead>
              <tbody id="studentHistoryBody">
                <!-- History Rows -->
              </tbody>
            </table>
          </div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">ปิดหน้าต่าง</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Modal 2: ดูรูปภาพขยาย -->
  <div class="modal fade" id="imageModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-lg">
      <div class="modal-content">
        <div class="modal-body text-center p-2">
          <img id="modalImage" src="" class="img-fluid rounded" alt="Full Image">
        </div>
      </div>
    </div>
  </div>

  <!-- Modal 3: แก้ไขข้อมูล -->
  <div class="modal fade" id="editModal" tabindex="-1" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title fw-bold">แก้ไขข้อมูลการบันทึก</h5>
          <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
        </div>
        <div class="modal-body">
          <form id="editForm">
            <input type="hidden" id="editId">
            <div class="mb-3">
              <label class="form-label">รหัสนักศึกษา</label>
              <input type="text" id="editStudentId" class="form-control" required>
            </div>
            <div class="mb-3">
              <label class="form-label">ชื่อ-นามสกุล</label>
              <input type="text" id="editFullName" class="form-control" required>
            </div>
            <div class="mb-3">
              <label class="form-label">ชื่อกิจกรรม</label>
              <input type="text" id="editActivityName" class="form-control" required>
            </div>
            <div class="mb-3">
              <label class="form-label">จำนวนชั่วโมง</label>
              <input type="number" id="editHours" class="form-control" step="0.5" required>
            </div>
            <div class="mb-3">
              <label class="form-label">สถานะ</label>
              <select id="editStatus" class="form-select">
                <option value="pending">รอตรวจสอบ</option>
                <option value="approved">อนุมัติ</option>
                <option value="rejected">ปฏิเสธ</option>
              </select>
            </div>
          </form>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">ยกเลิก</button>
          <button type="button" class="btn btn-primary" onclick="saveEdit()">บันทึกการเปลี่ยนแปลง</button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    let rawData = [];
    let statusChartInstance = null;
    let dailyChartInstance = null;

    document.addEventListener('DOMContentLoaded', () => {
      fetchData();
    });

    async function fetchData() {
      try {
        const res = await fetch('/api/admin/volunteers');
        rawData = await res.json();
        renderKPIs(rawData);
        renderCharts(rawData);
        renderTable(rawData);
      } catch (err) {
        console.error('Fetch data error:', err);
      }
    }

    function renderKPIs(data) {
      document.getElementById('kpi-total').innerText = data.length;
      document.getElementById('kpi-pending').innerText = data.filter(d => d.status === 'pending').length;
      document.getElementById('kpi-approved').innerText = data.filter(d => d.status === 'approved').length;
      
      const approvedHours = data
        .filter(d => d.status === 'approved')
        .reduce((sum, d) => sum + d.hours, 0);
      document.getElementById('kpi-hours').innerText = approvedHours;
    }

    function renderCharts(data) {
      // 1. Doughnut Chart (Status Distribution)
      const pendingCount = data.filter(d => d.status === 'pending').length;
      const approvedCount = data.filter(d => d.status === 'approved').length;
      const rejectedCount = data.filter(d => d.status === 'rejected').length;

      const ctxStatus = document.getElementById('statusChart').getContext('2d');
      if (statusChartInstance) statusChartInstance.destroy();

      statusChartInstance = new Chart(ctxStatus, {
        type: 'doughnut',
        data: {
          labels: ['รอตรวจสอบ', 'อนุมัติแล้ว', 'ปฏิเสธ'],
          datasets: [{
            data: [pendingCount, approvedCount, rejectedCount],
            backgroundColor: ['#ffc107', '#198754', '#dc3545'],
            borderWidth: 2,
            hoverOffset: 4
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom', labels: { boxWidth: 12, font: { family: 'Prompt' } } }
          },
          cutout: '70%'
        }
      });

      // 2. Bar Chart (Daily Statistics - Modern Gradient)
      const dailyMap = {};
      data.forEach(d => {
        const dateStr = new Date(d.createdAt).toLocaleDateString('th-TH', { day: '2-digit', month: '2-digit' });
        dailyMap[dateStr] = (dailyMap[dateStr] || 0) + 1;
      });

      const dates = Object.keys(dailyMap).reverse().slice(-7);
      const counts = dates.map(k => dailyMap[k]);

      const ctxDaily = document.getElementById('dailyChart').getContext('2d');
      if (dailyChartInstance) dailyChartInstance.destroy();

      const gradient = ctxDaily.createLinearGradient(0, 0, 0, 200);
      gradient.addColorStop(0, 'rgba(13, 110, 253, 0.85)');
      gradient.addColorStop(1, 'rgba(13, 110, 253, 0.15)');

      dailyChartInstance = new Chart(ctxDaily, {
        type: 'bar',
        data: {
          labels: dates,
          datasets: [{
            label: 'จำนวนการบันทึก (รายการ)',
            data: counts,
            backgroundColor: gradient,
            borderColor: '#0d6efd',
            borderWidth: 1.5,
            borderRadius: 6,
            borderSkipped: false
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          },
          scales: {
            y: { beginAtZero: true, ticks: { precision: 0, font: { family: 'Prompt' } }, grid: { borderDash: [4, 4] } },
            x: { ticks: { font: { family: 'Prompt' } }, grid: { display: false } }
          }
        }
      });
    }

    function renderTable(data) {
      const tbody = document.getElementById('tableBody');
      tbody.innerHTML = '';

      if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted py-4">ไม่พบข้อมูล</td></tr>';
        return;
      }

      data.forEach(item => {
        const dateStr = new Date(item.createdAt).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' });
        
        let statusBadge = '<span class="badge bg-warning text-dark">รอตรวจสอบ</span>';
        if (item.status === 'approved') statusBadge = '<span class="badge bg-success">อนุมัติแล้ว</span>';
        if (item.status === 'rejected') statusBadge = '<span class="badge bg-danger">ปฏิเสธ</span>';

        const tr = document.createElement('tr');
        tr.innerHTML = \`
          <td>
            <img src="\${item.imageUrl}" class="table-img" onclick="openImageModal('\${item.imageUrl}')" alt="กิจกรรม">
          </td>
          <td class="fw-semibold">\${item.studentId}</td>
          <td>
            <span class="student-link" onclick="openStudentModal('\${item.studentId}')">
              <i class="fa-solid fa-user-graduate me-1"></i>\${item.fullName}
            </span>
          </td>
          <td>\${item.activityName}</td>
          <td class="fw-bold text-primary">\${item.hours}</td>
          <td class="small text-muted">\${dateStr}</td>
          <td>\${statusBadge}</td>
          <td class="text-center">
            <button class="btn btn-sm btn-outline-success me-1" onclick="updateStatus('\${item._id}', 'approved')" title="อนุมัติ"><i class="fa-solid fa-check"></i></button>
            <button class="btn btn-sm btn-outline-danger me-1" onclick="updateStatus('\${item._id}', 'rejected')" title="ปฏิเสธ"><i class="fa-solid fa-xmark"></i></button>
            <button class="btn btn-sm btn-outline-primary me-1" onclick="openEditModal('\${item._id}')" title="แก้ไข"><i class="fa-solid fa-pen"></i></button>
            <button class="btn btn-sm btn-outline-secondary" onclick="deleteRecord('\${item._id}')" title="ลบ"><i class="fa-solid fa-trash"></i></button>
          </td>
        \`;
        tbody.appendChild(tr);
      });
    }

    function filterData() {
      const searchText = document.getElementById('searchInput').value.toLowerCase();
      const statusValue = document.getElementById('statusFilter').value;

      const filtered = rawData.filter(item => {
        const matchSearch = item.studentId.toLowerCase().includes(searchText) ||
                            item.fullName.toLowerCase().includes(searchText) ||
                            item.activityName.toLowerCase().includes(searchText);
        const matchStatus = statusValue === 'all' || item.status === statusValue;
        return matchSearch && matchStatus;
      });

      renderTable(filtered);
    }

    // ==========================================
    // NEW FUNCTION: Student Individual Detail Modal
    // ==========================================
    function openStudentModal(studentId) {
      const studentRecords = rawData.filter(d => d.studentId === studentId);
      if (studentRecords.length === 0) return;

      const studentInfo = studentRecords[0];
      
      document.getElementById('detailFullName').innerText = studentInfo.fullName;
      document.getElementById('detailStudentId').innerText = studentInfo.studentId;
      document.getElementById('detailAvatar').src = studentInfo.pictureUrl || 'https://via.placeholder.com/70?text=USER';

      const totalApprovedHours = studentRecords
        .filter(d => d.status === 'approved')
        .reduce((sum, d) => sum + d.hours, 0);
      
      document.getElementById('detailTotalHours').innerText = totalApprovedHours;

      const historyBody = document.getElementById('studentHistoryBody');
      historyBody.innerHTML = '';

      studentRecords.forEach(item => {
        const dateStr = new Date(item.createdAt).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' });
        
        let statusBadge = '<span class="badge bg-warning text-dark">รอตรวจสอบ</span>';
        if (item.status === 'approved') statusBadge = '<span class="badge bg-success">อนุมัติแล้ว</span>';
        if (item.status === 'rejected') statusBadge = '<span class="badge bg-danger">ปฏิเสธ</span>';

        const tr = document.createElement('tr');
        tr.innerHTML = \`
          <td class="text-center">
            <img src="\${item.imageUrl}" class="table-img" onclick="openImageModal('\${item.imageUrl}')" alt="กิจกรรม">
          </td>
          <td>\${item.activityName}</td>
          <td class="text-center fw-bold text-primary">\${item.hours}</td>
          <td class="small text-muted">\${dateStr}</td>
          <td class="text-center">\${statusBadge}</td>
        \`;
        historyBody.appendChild(tr);
      });

      const modal = new bootstrap.Modal(document.getElementById('studentDetailModal'));
      modal.show();
    }

    // Existing Helper Functions
    function openImageModal(url) {
      document.getElementById('modalImage').src = url;
      const modal = new bootstrap.Modal(document.getElementById('imageModal'));
      modal.show();
    }

    async function updateStatus(id, newStatus) {
      try {
        await fetch(\`/api/admin/volunteers/\${id}\`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ status: newStatus })
        });
        fetchData();
      } catch (err) {
        console.error('Update status error:', err);
      }
    }

    function openEditModal(id) {
      const item = rawData.find(d => d._id === id);
      if (!item) return;

      document.getElementById('editId').value = item._id;
      document.getElementById('editStudentId').value = item.studentId;
      document.getElementById('editFullName').value = item.fullName;
      document.getElementById('editActivityName').value = item.activityName;
      document.getElementById('editHours').value = item.hours;
      document.getElementById('editStatus').value = item.status;

      const modal = new bootstrap.Modal(document.getElementById('editModal'));
      modal.show();
    }

    async function saveEdit() {
      const id = document.getElementById('editId').value;
      const payload = {
        studentId: document.getElementById('editStudentId').value,
        fullName: document.getElementById('editFullName').value,
        activityName: document.getElementById('editActivityName').value,
        hours: parseFloat(document.getElementById('editHours').value),
        status: document.getElementById('editStatus').value
      };

      try {
        await fetch(\`/api/admin/volunteers/\${id}\`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        const modalEl = document.getElementById('editModal');
        const modal = bootstrap.Modal.getInstance(modalEl);
        modal.hide();

        fetchData();
      } catch (err) {
        console.error('Save edit error:', err);
      }
    }

    async function deleteRecord(id) {
      if (!confirm('คุณแน่ใจหรือไม่ที่จะลบรายการนี้?')) return;

      try {
        await fetch(\`/api/admin/volunteers/\${id}\`, { method: 'DELETE' });
        fetchData();
      } catch (err) {
        console.error('Delete error:', err);
      }
    }
  </script>
</body>
</html>
  `);
});


// ==========================================
// 7. SERVER START
// ==========================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
  console.log(`Admin Dashboard: http://localhost:${PORT}/admin`);
});