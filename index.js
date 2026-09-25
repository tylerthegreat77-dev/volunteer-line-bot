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
app.post('/webhook', line.middleware(config), async (req, res) => {
  try {
    console.log('--- WEBHOOK HIT ---', JSON.stringify(req.body.events, null, 2));

    // 1. Safety Guard: เช็กว่ามี events ส่งมาจริงไหม ถ้าไม่มีให้ตอบ 200 OK กลับไปเลย
    if (!req.body.events || !Array.isArray(req.body.events) || req.body.events.length === 0) {
      return res.status(200).send('OK');
    }

    // 2. รัน handleEvent
    const results = await Promise.all(req.body.events.map(handleEvent));
    res.status(200).json(results);
  } catch (err) {
    console.error('❌ Webhook Error:', err);
    // ตอบ 200 เพื่อไม่ให้ LINE retry ยิงซ้ำเวลาเกิดข้อผิดพลาดภายใน
    res.status(200).end();
  }
});

// Middleware สำหรับ Route อื่นๆ (ต้องอยู่ใต้ Webhook เท่านั้น)
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// เชื่อมต่อ MongoDB (ใส่ Fallback URI สำรองป้องกัน MONGODB_URI ใน Render อ่านค่าไม่ได้)
const MONGO_URI = process.env.MONGODB_URI || 'mongodb+srv://volunteer_user:wfvZwcF3XuRdvhZy@cluster0.3rc3oyf.mongodb.net/volunteer_db?appName=Cluster0';

mongoose.connect(MONGO_URI)
  .then(() => {
    console.log('✅ เชื่อมต่อ MongoDB สำเร็จแล้ว!');
    console.log('📌 DB Name ที่ใช้งานอยู่จริง:', mongoose.connection.name);
  })
  .catch((err) => console.error('❌ เชื่อมต่อ MongoDB ผิดพลาด:', err));

// Schema สำหรับเก็บข้อมูลจิตอาสา
const volunteerSchema = new mongoose.Schema({
  userId: String,
  facultyCode: String,
  facultyName: String,
  studentId: String,
  name: { type: String, default: 'ไม่ระบุชื่อ' },
  hours: Number,
  activityName: { type: String, default: 'ไม่ระบุกิจกรรม' },
  imageUrl: { type: String, default: '' },
  date: { type: Date, default: Date.now }
});

const Volunteer = mongoose.model('Volunteer', volunteerSchema, 'records');
// Schema สำหรับเก็บรูปภาพชั่วคราว
const tempImageSchema = new mongoose.Schema({
  userId: String,
  imageUrl: String,
  createdAt: { type: Date, default: Date.now, expires: 1800 }
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

// 2. API แก้ไขข้อมูลรายการบันทึก (เพิ่มใหม่)
app.put('/api/admin/records/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { facultyCode, studentId, name, hours, activityName } = req.body;

    const facultyName = FACULTY_MAP[facultyCode] || 'ไม่ระบุคณะ';

    const updatedRecord = await Volunteer.findByIdAndUpdate(
      id,
      {
        facultyCode,
        facultyName,
        studentId,
        name,
        hours: Number(hours),
        activityName
      },
      { new: true }
    );

    if (!updatedRecord) {
      return res.status(404).json({ success: false, message: 'ไม่พบรายการที่ต้องการแก้ไข' });
    }

    res.json({ success: true, message: 'อัปเดตข้อมูลสำเร็จ', data: updatedRecord });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 3. API ลบรายการบันทึก (เพิ่มใหม่)
app.delete('/api/admin/records/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const deletedRecord = await Volunteer.findByIdAndDelete(id);

    if (!deletedRecord) {
      return res.status(404).json({ success: false, message: 'ไม่พบรายการที่ต้องการลบ' });
    }

    res.json({ success: true, message: 'ลบข้อมูลสำเร็จ' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 4. Export Excel
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
      { header: 'ชื่อกิจกรรม', key: 'activityName', width: 30 },
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
        activityName: v.activityName || 'ไม่ระบุกิจกรรม',
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

// 5. หน้า Admin Dashboard สไตล์ Modern Pro UI + Visual Analytics + Complete CRUD
app.get('/admin', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Volunteer Pro Admin Dashboard</title>
      
      <!-- Typography & UI Frameworks -->
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
      <script src="https://unpkg.com/lucide@latest"></script>
      <!-- Chart.js -->
      <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

      <style>
        :root {
          --bg-main: #f8fafc;
          --border-color: #e2e8f0;
          --primary-color: #4f46e5;
          --primary-hover: #4338ca;
          --card-shadow: 0 10px 15px -3px rgba(15, 23, 42, 0.03), 0 4px 6px -4px rgba(15, 23, 42, 0.02);
        }

        body {
          font-family: 'Sarabun', 'Plus Jakarta Sans', sans-serif;
          background-color: var(--bg-main);
          color: #1e293b;
        }

        .navbar {
          background: rgba(255, 255, 255, 0.85);
          backdrop-filter: blur(12px);
          border-bottom: 1px solid var(--border-color);
        }

        .brand-icon {
          width: 40px;
          height: 40px;
          background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%);
          color: white;
          border-radius: 12px;
          display: flex;
          align-items: center;
          justify-content: center;
          box-shadow: 0 4px 12px rgba(79, 70, 229, 0.25);
        }

        .card {
          border: 1px solid var(--border-color);
          border-radius: 20px;
          box-shadow: var(--card-shadow);
          background: #ffffff;
          transition: transform 0.2s ease, box-shadow 0.2s ease;
        }

        .stat-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 12px 20px -5px rgba(15, 23, 42, 0.08);
        }

        .stat-icon {
          width: 52px;
          height: 52px;
          border-radius: 14px;
          display: flex;
          align-items: center;
          justify-content: center;
        }

        .table-card {
          overflow: hidden;
        }

        .table > :not(caption) > * > * {
          padding: 1rem 1.25rem;
          border-bottom-color: #f1f5f9;
        }

        .table thead th {
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.08em;
          color: #64748b;
          font-weight: 700;
          background-color: #f8fafc;
        }

        .img-thumb {
          width: 46px;
          height: 46px;
          object-fit: cover;
          border-radius: 12px;
          cursor: pointer;
          border: 1px solid #e2e8f0;
          transition: transform 0.2s ease;
        }

        .img-thumb:hover {
          transform: scale(1.1);
        }

        .badge-faculty {
          font-size: 0.75rem;
          padding: 0.4em 0.8em;
          border-radius: 8px;
          font-weight: 600;
        }

        .badge-hours {
          background-color: #ecfdf5;
          color: #047857;
          border: 1px solid #a7f3d0;
          font-weight: 700;
          padding: 0.4em 0.8em;
          border-radius: 30px;
        }

        .btn-action {
          width: 34px;
          height: 34px;
          padding: 0;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          border-radius: 8px;
        }

        .chart-container {
          position: relative;
          height: 260px;
          width: 100%;
        }

        .modal-content {
          border-radius: 20px;
          border: none;
          box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
        }
      </style>
    </head>
    <body>

      <!-- Navbar -->
      <nav class="navbar navbar-expand-lg sticky-top py-3">
        <div class="container-fluid px-4">
          <a class="navbar-brand d-flex align-items-center gap-3 fw-bold text-dark" href="#">
            <div class="brand-icon">
              <i data-lucide="heart-handshake" style="width:22px;"></i>
            </div>
            <div>
              <div class="fs-5 lh-1">Volunteer Admin</div>
              <span class="text-muted fw-normal" style="font-size:0.75rem;">ระบบจัดการชั่วโมงจิตอาสา</span>
            </div>
          </a>
          <div class="d-flex gap-2">
            <a href="/admin/export-excel" class="btn btn-outline-success d-flex align-items-center gap-2 px-3" style="border-radius:12px; font-weight:500;">
              <i data-lucide="file-spreadsheet" style="width:18px;"></i> Export Excel
            </a>
            <button class="btn btn-primary d-flex align-items-center gap-2 px-3" style="border-radius:12px; background:var(--primary-color); font-weight:500;" onclick="loadData()">
              <i data-lucide="refresh-cw" style="width:18px;"></i> รีเฟรช
            </button>
          </div>
        </div>
      </nav>

      <div class="container-fluid px-4 py-4">
        
        <!-- Summary Cards -->
        <div class="row g-3 mb-4">
          <div class="col-12 col-md-4">
            <div class="card stat-card p-3">
              <div class="d-flex align-items-center justify-content-between">
                <div>
                  <div class="text-muted small fw-semibold mb-1">จำนวนการบันทึกทั้งหมด</div>
                  <h2 class="fw-bold mb-0 text-dark" id="statCount">0</h2>
                </div>
                <div class="stat-icon bg-primary-subtle text-primary">
                  <i data-lucide="clipboard-list" style="width:26px; height:26px;"></i>
                </div>
              </div>
            </div>
          </div>
          <div class="col-12 col-md-4">
            <div class="card stat-card p-3">
              <div class="d-flex align-items-center justify-content-between">
                <div>
                  <div class="text-muted small fw-semibold mb-1">ชั่วโมงสะสมรวมทั้งหมด</div>
                  <h2 class="fw-bold mb-0 text-success" id="statHours">0 <small class="fs-6">ชม.</small></h2>
                </div>
                <div class="stat-icon bg-success-subtle text-success">
                  <i data-lucide="clock" style="width:26px; height:26px;"></i>
                </div>
              </div>
            </div>
          </div>
          <div class="col-12 col-md-4">
            <div class="card stat-card p-3">
              <div class="d-flex align-items-center justify-content-between">
                <div>
                  <div class="text-muted small fw-semibold mb-1">นักศึกษาที่เข้าร่วม</div>
                  <h2 class="fw-bold mb-0 text-indigo" id="statStudents">0 <small class="fs-6">คน</small></h2>
                </div>
                <div class="stat-icon bg-warning-subtle text-warning">
                  <i data-lucide="users" style="width:26px; height:26px;"></i>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- Charts Section -->
        <div class="row g-3 mb-4">
          <div class="col-12 col-lg-7">
            <div class="card p-3">
              <h6 class="fw-bold text-dark mb-3 d-flex align-items-center gap-2">
                <i data-lucide="bar-chart-3" class="text-primary" style="width:18px;"></i> สรุปชั่วโมงจำแนกตามคณะ
              </h6>
              <div class="chart-container">
                <canvas id="facultyChart"></canvas>
              </div>
            </div>
          </div>
          <div class="col-12 col-lg-5">
            <div class="card p-3">
              <h6 class="fw-bold text-dark mb-3 d-flex align-items-center gap-2">
                <i data-lucide="pie-chart" class="text-primary" style="width:18px;"></i> สัดส่วนจำนวนครั้งการเข้าร่วม
              </h6>
              <div class="chart-container">
                <canvas id="ratioChart"></canvas>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Box -->
        <div class="card p-3 mb-4">
          <div class="row g-3">
            <div class="col-12 col-md-8">
              <div class="input-group">
                <span class="input-group-text bg-white border-end-0 pe-0" style="border-radius: 12px 0 0 12px; border-color: #cbd5e1;">
                  <i data-lucide="search" class="text-muted" style="width:18px;"></i>
                </span>
                <input type="text" id="searchInput" class="form-control border-start-0" style="border-radius: 0 12px 12px 0; border-color: #cbd5e1;" placeholder="ค้นหาด้วย รหัสนักศึกษา, ชื่อ-นามสกุล หรือกิจกรรม..." onkeyup="filterTable()">
              </div>
            </div>
            <div class="col-12 col-md-4">
              <select id="facultyFilter" class="form-select" style="border-radius: 12px; border-color: #cbd5e1;" onchange="filterTable()">
                <option value="">🏛️ แสดงทุกคณะ</option>
                <option value="01">01 - คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร</option>
                <option value="02">02 - คณะบริหารธุรกิจและศิลปศาสตร์</option>
                <option value="03">03 - คณะวิศวกรรมศาสตร์</option>
              </select>
            </div>
          </div>
        </div>

        <!-- Data Table -->
        <div class="card table-card">
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
          <div class="modal-content p-2">
            <div class="modal-header border-0 pb-0">
              <h5 class="modal-header-title fw-bold">✏️ แก้ไขข้อมูลบันทึกจิตอาสา</h5>
              <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body">
              <form id="editForm">
                <input type="hidden" id="editId">
                <div class="mb-3">
                  <label class="form-label text-muted small fw-bold">คณะ</label>
                  <select id="editFacultyCode" class="form-select" required style="border-radius:10px;">
                    <option value="01">01 - คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร</option>
                    <option value="02">02 - คณะบริหารธุรกิจและศิลปศาสตร์</option>
                    <option value="03">03 - คณะวิศวกรรมศาสตร์</option>
                  </select>
                </div>
                <div class="mb-3">
                  <label class="form-label text-muted small fw-bold">รหัสนักศึกษา</label>
                  <input type="text" id="editStudentId" class="form-control" required style="border-radius:10px;">
                </div>
                <div class="mb-3">
                  <label class="form-label text-muted small fw-bold">ชื่อ-นามสกุล</label>
                  <input type="text" id="editName" class="form-control" required style="border-radius:10px;">
                </div>
                <div class="mb-3">
                  <label class="form-label text-muted small fw-bold">จำนวนชั่วโมง</label>
                  <input type="number" id="editHours" class="form-control" min="0.5" step="0.5" required style="border-radius:10px;">
                </div>
                <div class="mb-3">
                  <label class="form-label text-muted small fw-bold">ชื่อกิจกรรม</label>
                  <input type="text" id="editActivityName" class="form-control" required style="border-radius:10px;">
                </div>
                <div class="d-flex justify-content-end gap-2 pt-2">
                  <button type="button" class="btn btn-light" data-bs-dismiss="modal" style="border-radius:10px;">ยกเลิก</button>
                  <button type="submit" class="btn btn-primary" style="border-radius:10px; background:var(--primary-color);">บันทึกการเปลี่ยนแปลง</button>
                </div>
              </form>
            </div>
          </div>
        </div>
      </div>

      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
      <script>
        let allRecords = [];
        let facultyChart = null;
        let ratioChart = null;

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
          const facultyHours = { '01': 0, '02': 0, '03': 0 };
          const facultyCounts = { '01': 0, '02': 0, '03': 0 };

          data.forEach(item => {
            if (facultyHours[item.facultyCode] !== undefined) {
              facultyHours[item.facultyCode] += (item.hours || 0);
              facultyCounts[item.facultyCode] += 1;
            }
          });

          // 1. Bar Chart (ชั่วโมงรวมแต่ละคณะ)
          const ctxBar = document.getElementById('facultyChart').getContext('2d');
          if (facultyChart) facultyChart.destroy();
          facultyChart = new Chart(ctxBar, {
            type: 'bar',
            data: {
              labels: ['วิทยาศาสตร์ฯ (01)', 'บริหารธุรกิจฯ (02)', 'วิศวกรรมศาสตร์ (03)'],
              datasets: [{
                label: 'ชั่วโมงสะสม',
                data: [facultyHours['01'], facultyHours['02'], facultyHours['03']],
                backgroundColor: ['rgba(16, 185, 129, 0.85)', 'rgba(79, 70, 229, 0.85)', 'rgba(245, 158, 11, 0.85)'],
                borderRadius: 8
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { display: false } },
              scales: { y: { beginAtZero: true } }
            }
          });

          // 2. Doughnut Chart (สัดส่วนจำนวนครั้ง)
          const ctxDoughnut = document.getElementById('ratioChart').getContext('2d');
          if (ratioChart) ratioChart.destroy();
          ratioChart = new Chart(ctxDoughnut, {
            type: 'doughnut',
            data: {
              labels: ['คณะ 01', 'คณะ 02', 'คณะ 03'],
              datasets: [{
                data: [facultyCounts['01'], facultyCounts['02'], facultyCounts['03']],
                backgroundColor: ['#10b981', '#4f46e5', '#f59e0b']
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { position: 'bottom' } }
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
              ? \`<img src="\${item.imageUrl}" class="img-thumb shadow-sm" onclick="showModal('\${item.imageUrl}')">\`
              : \`<span class="badge bg-light text-muted border py-2 px-2" style="font-size:11px;">ไม่มีรูป</span>\`;

            const facultyClass = facultyColors[item.facultyCode] || 'bg-light text-dark';
            const facultyBadge = item.facultyCode 
              ? \`<span class="badge badge-faculty border \${facultyClass}">[\${item.facultyCode}] \${item.facultyName || ''}</span>\`
              : \`<span class="text-muted small">ไม่ระบุ</span>\`;

            const tr = document.createElement('tr');
            tr.innerHTML = \`
              <td>\${imgHtml}</td>
              <td>
                <div class="fw-bold text-dark">\${item.name || 'ไม่ระบุชื่อ'}</div>
                <div class="text-muted small">🆔 \${item.studentId}</div>
              </td>
              <td>\${facultyBadge}</td>
              <td><div class="fw-medium text-dark">\${item.activityName || 'ไม่ระบุกิจกรรม'}</div></td>
              <td><span class="badge badge-hours">+\${item.hours} ชม.</span></td>
              <td>
                <div class="small fw-medium text-dark">\${dateStr}</div>
                <div class="text-muted" style="font-size: 11px;">\${timeStr} น.</div>
              </td>
              <td class="text-end">
                <button class="btn btn-action btn-outline-primary me-1" onclick="openEditModal('\${item._id}')" title="แก้ไข">
                  <i data-lucide="edit-3" style="width:16px;"></i>
                </button>
                <button class="btn btn-action btn-outline-danger" onclick="deleteRecord('\${item._id}')" title="ลบ">
                  <i data-lucide="trash-2" style="width:16px;"></i>
                </button>
              </td>
            \`;
            tbody.appendChild(tr);
          });

          lucide.createIcons();
        }

        function showModal(url) {
          document.getElementById('modalImg').src = url;
          new bootstrap.Modal(document.getElementById('imageModal')).show();
        }

        function openEditModal(id) {
          const item = allRecords.find(r => r._id === id);
          if (!item) return;

          document.getElementById('editId').value = item._id;
          document.getElementById('editFacultyCode').value = item.facultyCode || '01';
          document.getElementById('editStudentId').value = item.studentId || '';
          document.getElementById('editName').value = item.name || '';
          document.getElementById('editHours').value = item.hours || 0;
          document.getElementById('editActivityName').value = item.activityName || '';

          new bootstrap.Modal(document.getElementById('editModal')).show();
        }

        document.getElementById('editForm').addEventListener('submit', async (e) => {
          e.preventDefault();
          const id = document.getElementById('editId').value;
          const payload = {
            facultyCode: document.getElementById('editFacultyCode').value,
            studentId: document.getElementById('editStudentId').value,
            name: document.getElementById('editName').value,
            hours: document.getElementById('editHours').value,
            activityName: document.getElementById('editActivityName').value
          };

          try {
            const res = await fetch(\`/api/admin/records/\${id}\`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(payload)
            });
            const result = await res.json();
            if (result.success) {
              bootstrap.Modal.getInstance(document.getElementById('editModal')).hide();
              loadData();
            } else {
              alert('แก้ไขไม่สำเร็จ: ' + result.message);
            }
          } catch (err) {
            alert('เกิดข้อผิดพลาดในการส่งข้อมูล');
          }
        });

        async function deleteRecord(id) {
          if (!confirm('คุณแน่ใจหรือไม่ว่าต้องการลบรายการนี้?')) return;

          try {
            const res = await fetch(\`/api/admin/records/\${id}\`, { method: 'DELETE' });
            const result = await res.json();
            if (result.success) {
              loadData();
            } else {
              alert('ลบไม่สำเร็จ: ' + result.message);
            }
          } catch (err) {
            alert('เกิดข้อผิดพลาดในการลบข้อมูล');
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

// ==========================================
// 💬 LINE BOT HANDLER
// ==========================================

// ✅ แก้ไขให้เป็นรูปแบบมาตรฐานของ LINE SDK v7.x
function replyTextMsg(replyToken, text) {
  return client.replyMessage(replyToken, {
    type: 'text',
    text: text
  });
}

async function handleImageMessage(event) {
  const userId = event.source.userId;
  const messageId = event.message.id;

  try {
    const stream = await blobClient.getMessageContent(messageId);
    
    const chunks = [];
    for await (const chunk of stream) {
      chunks.push(chunk);
    }
    const buffer = Buffer.concat(chunks);
    const base64Image = `data:image/jpeg;base64,${buffer.toString('base64')}`;

    const uploadResult = await cloudinary.uploader.upload(base64Image, {
      folder: 'volunteer_proofs'
    });

    const imageUrl = uploadResult.secure_url;

    await TempImage.findOneAndUpdate(
      { userId },
      { imageUrl, createdAt: new Date() },
      { upsert: true, new: true }
    );

    return replyTextMsg(
      event.replyToken,
      '📷 ได้รับรูปภาพหลักฐานเรียบร้อยแล้วครับ!\n\nกรุณาพิมพ์บันทึกชั่วโมงต่อได้เลย เช่น:\nบันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด'
    );

  } catch (error) {
    console.error('Error handling image with Cloudinary:', error);
    return replyTextMsg(event.replyToken, '❌ ไม่สามารถบันทึกรูปภาพได้ กรุณาลองส่งใหม่อีกครั้ง');
  }
}

async function handleEvent(event) {
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

    if (parts.length < 6) {
      return replyTextMsg(
        event.replyToken, 
        '❌ รูปแบบคำสั่งไม่ถูกต้อง!\nกรุณาพิมพ์: บันทึก <รหัสคณะ> <รหัสนักศึกษา> <ชื่อ-นามสกุล> <จำนวนชั่วโมง> <ชื่อกิจกรรม>\n\n' +
        '🏛 รหัสคณะ:\n01 = วิทยาศาสตร์และเทคโนโลยีการเกษตร\n02 = บริหารธุรกิจและศิลปศาสตร์\n03 = วิศวกรรมศาสตร์\n\n' +
        'ตัวอย่าง:\nบันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด'
      );
    }

    const facultyCode = parts[1];
    const studentId = parts[2];

    let hoursIndex = -1;
    for (let i = 3; i < parts.length - 1; i++) {
      if (!isNaN(parts[i])) {
        hoursIndex = i;
        break;
      }
    }

    if (hoursIndex === -1) {
      return replyTextMsg(
        event.replyToken, 
        '❌ ไม่พบจำนวนชั่วโมงที่เป็นตัวเลข หรือพิมพ์รูปแบบไม่ถูกต้อง\nตัวอย่างที่ถูกต้อง:\nบันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด'
      );
    }

    const name = parts.slice(3, hoursIndex).join(' ').trim();
    const hours = parseFloat(parts[hoursIndex]);
    const activityName = parts.slice(hoursIndex + 1).join(' ').trim();

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

    if (!activityName) {
      return replyTextMsg(event.replyToken, '❌ กรุณาระบุชื่อกิจกรรมต่อท้ายด้วยครับ');
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
        activityName, 
        imageUrl 
      });

      await newRecord.save();

      const allRecords = await Volunteer.find({ studentId });
      const totalHours = allRecords.reduce((sum, item) => sum + item.hours, 0);

      let replyText = `✅ บันทึกชั่วโมงจิตอาสาสำเร็จ!\n\n` +
                        `👤 ชื่อ: ${name}\n` +
                        `🆔 รหัส: ${studentId}\n` +
                        `🏛 คณะ: ${facultyName}\n` +
                        `📌 กิจกรรม: ${activityName}\n` +
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
                   `บันทึก <รหัสคณะ> <รหัสประจำตัว> <ชื่อ-นามสกุล> <ชั่วโมง> <ชื่อกิจกรรม>\n\n` +
                   `🏛 รหัสคณะ:\n` +
                   `01 = คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร\n` +
                   `02 = คณะบริหารธุรกิจและศิลปศาสตร์\n` +
                   `03 = คณะวิศวกรรมศาสตร์\n\n` +
                   `💡 ตัวอย่าง:\nบันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด\n\n` +
                   `🔍 เช็คชั่วโมงสะสม:\n` +
                   `พิมพ์: เช็คชั่วโมง <รหัสนักศึกษา>`;

  return replyTextMsg(event.replyToken, helpText);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server เปิดทำงานแล้วที่ Port ${PORT}`);
});