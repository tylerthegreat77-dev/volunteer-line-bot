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

// Messaging API Client
const client = line.messagingApi 
  ? new line.messagingApi.MessagingApiClient({ channelAccessToken: config.channelAccessToken })
  : new line.Client(config);

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

const Volunteer = mongoose.model('Volunteer', volunteerSchema);

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

// 2. Export Excel
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

// 3. หน้า Admin Dashboard (ดีไซน์ถอดแบบจากภาพที่ส่งมา)
app.get('/admin', (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Volunteer Dashboard</title>
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
      <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
      <script src="https://unpkg.com/lucide@latest"></script>
      <style>
        :root {
          --sidebar-bg: #636e88;
          --sidebar-hover: #525c74;
          --main-bg: #e2e7f0;
          --card-bg: #ffffff;
          --text-muted: #8c98a9;
          --accent-blue: #0088ff;
          --accent-orange: #ffaa00;
          --accent-green: #00cc88;
          --accent-purple: #aa00ff;
        }

        body {
          font-family: 'Sarabun', 'Plus Jakarta Sans', sans-serif;
          background-color: var(--main-bg);
          margin: 0;
          padding: 0;
          height: 100vh;
          overflow-x: hidden;
        }

        .dashboard-container {
          display: flex;
          min-height: 100vh;
        }

        /* Sidebar Styling */
        .sidebar {
          width: 240px;
          background-color: var(--sidebar-bg);
          color: #ffffff;
          display: flex;
          flex-direction: column;
          padding: 24px 16px;
          flex-shrink: 0;
        }

        .sidebar-brand {
          font-weight: 700;
          font-size: 1.25rem;
          letter-spacing: 1px;
          margin-bottom: 32px;
          padding-left: 12px;
        }

        .nav-list {
          list-style: none;
          padding: 0;
          margin: 0;
        }

        .nav-item {
          margin-bottom: 8px;
        }

        .nav-link {
          color: #d1d5db;
          display: flex;
          align-items: center;
          gap: 14px;
          padding: 10px 14px;
          border-radius: 8px;
          text-decoration: none;
          font-size: 0.95rem;
          transition: all 0.2s ease;
        }

        .nav-link:hover, .nav-link.active {
          color: #ffffff;
          background-color: var(--sidebar-hover);
        }

        .nav-link.active {
          color: #38bdf8;
        }

        .menu-divider {
          font-size: 0.75rem;
          color: #9ca3af;
          margin: 24px 0 12px 12px;
          text-transform: uppercase;
        }

        /* Main Content */
        .main-wrapper {
          flex: 1;
          display: flex;
          flex-direction: column;
          overflow-y: auto;
        }

        /* Top Bar */
        .top-bar {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 16px 32px;
          background-color: transparent;
        }

        .search-container {
          position: relative;
          width: 280px;
        }

        .search-container input {
          width: 100%;
          padding: 8px 16px 8px 36px;
          border-radius: 20px;
          border: 1px solid #cbd5e1;
          background: #ffffff;
          outline: none;
          font-size: 0.85rem;
        }

        .search-container i {
          position: absolute;
          left: 12px;
          top: 50%;
          transform: translateY(-50%);
          color: #94a3b8;
          width: 16px;
        }

        .user-profile {
          display: flex;
          align-items: center;
          gap: 10px;
          color: #64748b;
          font-size: 0.85rem;
        }

        .avatar {
          width: 32px;
          height: 32px;
          border-radius: 50%;
          background: #cbd5e1;
          display: flex;
          align-items: center;
          justify-content: center;
        }

        /* Dashboard Grid Layout */
        .dashboard-content {
          padding: 0 32px 32px 32px;
        }

        .card-custom {
          background: var(--card-bg);
          border-radius: 12px;
          padding: 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          height: 100%;
        }

        .card-title-custom {
          font-size: 0.8rem;
          color: var(--text-muted);
          font-weight: 600;
          margin-bottom: 12px;
        }

        /* Stat Mini Cards */
        .stat-value {
          font-size: 1.5rem;
          font-weight: 700;
        }

        .mini-chart-container {
          height: 45px;
          margin-top: 8px;
        }

        /* Table Area */
        .table-custom {
          width: 100%;
          font-size: 0.85rem;
        }

        .table-custom th {
          color: #94a3b8;
          font-weight: 600;
          border-bottom: 1px solid #f1f5f9;
          padding-bottom: 8px;
        }

        .table-custom td {
          padding: 10px 0;
          border-bottom: 1px solid #f8fafc;
        }

        .img-thumb {
          width: 36px;
          height: 36px;
          object-fit: cover;
          border-radius: 6px;
          cursor: pointer;
        }

        /* Custom Progress Bars for Right Side */
        .progress-item {
          margin-bottom: 14px;
        }

        .progress-label {
          display: flex;
          justify-content: space-between;
          font-size: 0.75rem;
          color: #64748b;
          margin-bottom: 4px;
        }

        .progress-bar-custom {
          height: 6px;
          border-radius: 3px;
        }
      </style>
    </head>
    <body>

      <div class="dashboard-container">
        
        <!-- Sidebar ด้านซ้าย -->
        <aside class="sidebar">
          <div class="sidebar-brand">DASHBOARD</div>
          <ul class="nav-list">
            <li class="nav-item">
              <a href="#" class="nav-link active">
                <i data-lucide="home" style="width:18px;"></i> Home
              </a>
            </li>
            <li class="nav-item">
              <a href="#" class="nav-link">
                <i data-lucide="bar-chart-2" style="width:18px;"></i> Charts
              </a>
            </li>
            <li class="nav-item">
              <a href="#" class="nav-link">
                <i data-lucide="star" style="width:18px;"></i> Favorites
              </a>
            </li>
            <li class="nav-item">
              <a href="#" class="nav-link">
                <i data-lucide="message-square" style="width:18px;"></i> Chat
              </a>
            </li>
            <li class="nav-item">
              <a href="#" class="nav-link">
                <i data-lucide="settings" style="width:18px;"></i> Setting
              </a>
            </li>
            <li class="nav-item">
              <a href="#" class="nav-link">
                <i data-lucide="help-circle" style="width:18px;"></i> Help
              </a>
            </li>
          </ul>

          <div class="menu-divider">Dashboard Menu</div>
          <ul class="nav-list">
            <li class="nav-item">
              <a href="/admin/export-excel" class="nav-link">
                <i data-lucide="file-spreadsheet" style="width:18px;"></i> Export Excel
              </a>
            </li>
            <li class="nav-item">
              <a href="#" class="nav-link" onclick="loadData(); return false;">
                <i data-lucide="refresh-cw" style="width:18px;"></i> Refresh Data
              </a>
            </li>
          </ul>
        </aside>

        <!-- Main Content Area -->
        <main class="main-wrapper">
          
          <!-- Top Bar -->
          <div class="top-bar">
            <div class="search-container">
              <i data-lucide="search"></i>
              <input type="text" id="searchInput" placeholder="ค้นหา รหัสนักศึกษา, ชื่อ..." onkeyup="filterData()">
            </div>
            <div class="user-profile">
              <span>hi, Admin</span>
              <div class="avatar"><i data-lucide="user" style="width:18px;"></i></div>
            </div>
          </div>

          <!-- Dashboard Content Grid -->
          <div class="dashboard-content">
            <div class="row g-3">
              
              <!-- Left Grid (Center Layout ในภาพ) -->
              <div class="col-12 col-lg-9">
                <div class="row g-3">
                  
                  <!-- Top Line Chart (Detailed Chart 01) -->
                  <div class="col-12">
                    <div class="card-custom">
                      <div class="card-title-custom">Detailed Chart 01 (แนวโน้มการบันทึกชั่วโมง)</div>
                      <div style="height: 180px;">
                        <canvas id="mainLineChart"></canvas>
                      </div>
                    </div>
                  </div>

                  <!-- 3 Mini Stat Cards -->
                  <div class="col-12 col-md-4">
                    <div class="card-custom">
                      <div class="card-title-custom">ชั่วโมงสะสมรวม</div>
                      <div class="stat-value text-primary" id="totalHoursText">0</div>
                      <div class="mini-chart-container">
                        <canvas id="miniChart1"></canvas>
                      </div>
                    </div>
                  </div>
                  
                  <div class="col-12 col-md-4">
                    <div class="card-custom">
                      <div class="card-title-custom">จำนวนการบันทึก</div>
                      <div class="stat-value text-warning" id="totalCountText">0</div>
                      <div class="mini-chart-container">
                        <canvas id="miniChart2"></canvas>
                      </div>
                    </div>
                  </div>

                  <div class="col-12 col-md-4">
                    <div class="card-custom">
                      <div class="card-title-custom">นักศึกษาเข้าร่วม</div>
                      <div class="stat-value text-success" id="totalStudentsText">0</div>
                      <div class="mini-chart-container">
                        <canvas id="miniChart3"></canvas>
                      </div>
                    </div>
                  </div>

                  <!-- Donut Chart & Bar Chart -->
                  <div class="col-12 col-md-4">
                    <div class="card-custom text-center">
                      <div class="card-title-custom text-start">Profile Strength (ภาพรวมเป้าหมาย)</div>
                      <div style="height: 140px; position: relative;" class="d-flex align-items-center justify-content-center">
                        <canvas id="donutChart"></canvas>
                        <div style="position: absolute; font-weight:700; font-size:1.2rem;" id="donutPercentText">75%</div>
                      </div>
                    </div>
                  </div>

                  <div class="col-12 col-md-8">
                    <div class="card-custom">
                      <div class="card-title-custom">Detailed Chart 02 (สถิติตามคณะ)</div>
                      <div style="height: 140px;">
                        <canvas id="barChart"></canvas>
                      </div>
                    </div>
                  </div>

                </div>
              </div>

              <!-- Right Panel Grid -->
              <div class="col-12 col-lg-3">
                <div class="row g-3">
                  
                  <!-- Average Charts -->
                  <div class="col-12">
                    <div class="card-custom">
                      <div class="card-title-custom">Average Charts</div>
                      <div class="progress-item">
                        <div class="progress-label"><span>คณะวิทยาศาสตร์ฯ (01)</span><span>70%</span></div>
                        <div class="progress"><div class="progress-bar bg-danger progress-bar-custom" style="width: 70%"></div></div>
                      </div>
                      <div class="progress-item">
                        <div class="progress-label"><span>คณะบริหารธุรกิจฯ (02)</span><span>45%</span></div>
                        <div class="progress"><div class="progress-bar bg-primary progress-bar-custom" style="width: 45%"></div></div>
                      </div>
                      <div class="progress-item">
                        <div class="progress-label"><span>คณะวิศวกรรมศาสตร์ (03)</span><span>85%</span></div>
                        <div class="progress"><div class="progress-bar bg-warning progress-bar-custom" style="width: 85%"></div></div>
                      </div>
                      <div class="progress-item">
                        <div class="progress-label"><span>อื่นๆ</span><span>30%</span></div>
                        <div class="progress"><div class="progress-bar bg-success progress-bar-custom" style="width: 30%"></div></div>
                      </div>
                    </div>
                  </div>

                  <!-- Recently Activity / Record List -->
                  <div class="col-12">
                    <div class="card-custom" style="max-height: 300px; overflow-y: auto;">
                      <div class="card-title-custom">Recently Activity</div>
                      <table class="table-custom">
                        <thead>
                          <tr>
                            <th>รูป</th>
                            <th>ข้อมูล</th>
                            <th>ชม.</th>
                          </tr>
                        </thead>
                        <tbody id="recentTbody">
                          <tr><td colspan="3" class="text-center text-muted">กำลังโหลด...</td></tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>
              </div>

            </div>
          </div>

        </main>
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

      <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
      <script>
        let allRecords = [];
        let charts = {};

        async function loadData() {
          try {
            const res = await fetch('/api/admin/records');
            const result = await res.json();
            if (result.success) {
              allRecords = result.data || [];
              updateDashboard(allRecords);
            }
          } catch (err) {
            console.error('Error fetching data:', err);
          }
        }

        function updateDashboard(data) {
          // Stat Counters
          const totalHours = data.reduce((sum, item) => sum + (item.hours || 0), 0);
          const totalCount = data.length;
          const uniqueStudents = new Set(data.map(item => item.studentId)).size;

          document.getElementById('totalHoursText').innerText = totalHours.toLocaleString();
          document.getElementById('totalCountText').innerText = totalCount.toLocaleString();
          document.getElementById('totalStudentsText').innerText = uniqueStudents.toLocaleString();

          renderRecentTable(data);
          initCharts(data);
        }

        function renderRecentTable(data) {
          const tbody = document.getElementById('recentTbody');
          tbody.innerHTML = '';

          if (!data || data.length === 0) {
            tbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">ไม่มีข้อมูล</td></tr>';
            return;
          }

          data.slice(0, 10).forEach(item => {
            const imgHtml = item.imageUrl 
              ? '<img src="' + item.imageUrl + '" class="img-thumb" onclick="showModal(\'' + item.imageUrl + '\')">'
              : '<span class="text-muted" style="font-size:10px;">ไม่มีรูป</span>';

            const tr = document.createElement('tr');
            tr.innerHTML = '<td>' + imgHtml + '</td>' +
              '<td>' +
                '<div style="font-weight:600; font-size:0.8rem;">' + (item.name || 'ไม่ระบุชื่อ') + '</div>' +
                '<div style="font-size:0.7rem; color:#94a3b8;">' + (item.studentId || '-') + '</div>' +
              '</td>' +
              '<td><span class="badge bg-success-subtle text-success">+' + (item.hours || 0) + '</span></td>';
            tbody.appendChild(tr);
          });
        }

        function initCharts(data) {
          // Main Line Chart
          if (charts.mainLine) charts.mainLine.destroy();
          const ctxLine = document.getElementById('mainLineChart').getContext('2d');
          charts.mainLine = new Chart(ctxLine, {
            type: 'line',
            data: {
              labels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul'],
              datasets: [{
                label: 'ชั่วโมงจิตอาสา',
                data: [12, 19, 8, 25, 22, 30, 28],
                borderColor: '#ffaa00',
                backgroundColor: 'rgba(255,170,0,0.1)',
                tension: 0.4,
                pointBackgroundColor: ['#0088ff', '#0088ff', '#0088ff', '#0088ff', '#0088ff', '#0088ff', '#0088ff']
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { display: false } },
              scales: { y: { display: false }, x: { grid: { display: false } } }
            }
          });

          // Mini Area Chart 1
          createMiniChart('miniChart1', [10, 20, 15, 30, 25, 40], '#0088ff');
          // Mini Area Chart 2
          createMiniChart('miniChart2', [5, 15, 25, 20, 35, 30], '#ffaa00');
          // Mini Area Chart 3
          createMiniChart('miniChart3', [12, 18, 14, 22, 28, 35], '#00cc88');

          // Donut Chart
          if (charts.donut) charts.donut.destroy();
          const ctxDonut = document.getElementById('donutChart').getContext('2d');
          charts.donut = new Chart(ctxDonut, {
            type: 'doughnut',
            data: {
              datasets: [{
                data: [75, 25],
                backgroundColor: ['#ffaa00', '#e2e8f0'],
                borderWidth: 0
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              cutout: '80%',
              plugins: { legend: { display: false } }
            }
          });

          // Bar Chart
          if (charts.bar) charts.bar.destroy();
          const ctxBar = document.getElementById('barChart').getContext('2d');
          charts.bar = new Chart(ctxBar, {
            type: 'bar',
            data: {
              labels: ['01', '02', '03', 'อื่นๆ'],
              datasets: [{
                data: [12, 19, 15, 8],
                backgroundColor: ['#ff4d4d', '#0088ff', '#ffaa00', '#aa00ff'],
                borderRadius: 4
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { display: false } },
              scales: { y: { display: false }, x: { grid: { display: false } } }
            }
          });
        }

        function createMiniChart(id, data, color) {
          if (charts[id]) charts[id].destroy();
          const ctx = document.getElementById(id).getContext('2d');
          charts[id] = new Chart(ctx, {
            type: 'line',
            data: {
              labels: data.map((_, i) => i),
              datasets: [{
                data: data,
                borderColor: color,
                backgroundColor: color + '33',
                fill: true,
                tension: 0.4,
                pointRadius: 0
              }]
            },
            options: {
              responsive: true,
              maintainAspectRatio: false,
              plugins: { legend: { display: false } },
              scales: { y: { display: false }, x: { display: false } }
            }
          });
        }

        function showModal(url) {
          document.getElementById('modalImg').src = url;
          new bootstrap.Modal(document.getElementById('imageModal')).show();
        }

        function filterData() {
          const searchText = document.getElementById('searchInput').value.toLowerCase();
          const filtered = allRecords.filter(item => 
            (item.studentId && item.studentId.toLowerCase().includes(searchText)) ||
            (item.name && item.name.toLowerCase().includes(searchText))
          );
          renderRecentTable(filtered);
        }

        window.onload = function() {
          lucide.createIcons();
          loadData();
        };
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