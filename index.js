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
app.use(express.urlencoded({ extended: true }));
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

// 2. API ดูข้อมูลนักศึกษารายบุคคล
app.get('/api/admin/students/:studentId', async (req, res) => {
  try {
    const { studentId } = req.params;
    const records = await Volunteer.find({ studentId }).sort({ date: -1 });

    if (!records.length) {
      return res.status(404).json({
        success: false,
        message: 'ไม่พบข้อมูลนักศึกษา'
      });
    }

    const totalHours = records.reduce(
      (sum, item) => sum + (Number(item.hours) || 0),
      0
    );

    const latest =
      records.find(r => r.name && r.name !== 'ไม่ระบุชื่อ') ||
      records[0];

    res.json({
      success: true,
      data: {
        studentId,
        name: latest.name || 'ไม่ระบุชื่อ',
        facultyCode: latest.facultyCode || '',
        facultyName: latest.facultyName || 'ไม่ระบุคณะ',
        totalHours,
        totalRecords: records.length,
        records
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// 2. API แก้ไขข้อมูลรายการบันทึก
app.put('/api/admin/records/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const {
      facultyCode,
      studentId,
      name,
      hours,
      activityName
    } = req.body;

    const facultyName =
      FACULTY_MAP[facultyCode] || 'ไม่ระบุคณะ';

    const updatedRecord =
      await Volunteer.findByIdAndUpdate(
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
      return res.status(404).json({
        success: false,
        message: 'ไม่พบรายการที่ต้องการแก้ไข'
      });
    }

    res.json({
      success: true,
      message: 'อัปเดตข้อมูลสำเร็จ',
      data: updatedRecord
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// 3. API ลบรายการบันทึก
app.delete('/api/admin/records/:id', async (req, res) => {
  try {
    const { id } = req.params;

    const deletedRecord =
      await Volunteer.findByIdAndDelete(id);

    if (!deletedRecord) {
      return res.status(404).json({
        success: false,
        message: 'ไม่พบรายการที่ต้องการลบ'
      });
    }

    res.json({
      success: true,
      message: 'ลบข้อมูลสำเร็จ'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// 4. Export Excel
app.get('/admin/export-excel', async (req, res) => {
  try {
    const records =
      await Volunteer.find().sort({ date: -1 });

    const workbook = new ExcelJS.Workbook();
    const worksheet =
      workbook.addWorksheet('รายงานชั่วโมงจิตอาสา');

    worksheet.columns = [
      {
        header: 'รหัสคณะ',
        key: 'facultyCode',
        width: 12
      },
      {
        header: 'ชื่อคณะ',
        key: 'facultyName',
        width: 35
      },
      {
        header: 'รหัสนักศึกษา',
        key: 'studentId',
        width: 20
      },
      {
        header: 'ชื่อ-นามสกุล',
        key: 'name',
        width: 25
      },
      {
        header: 'ชื่อกิจกรรม',
        key: 'activityName',
        width: 30
      },
      {
        header: 'ชั่วโมงที่บันทึก',
        key: 'hours',
        width: 15
      },
      {
        header: 'วันที่บันทึก',
        key: 'date',
        width: 22
      },
      {
        header: 'ลิงก์รูปภาพหลักฐาน',
        key: 'imageUrl',
        width: 45
      }
    ];

    worksheet.getRow(1).font = {
      bold: true
    };

    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {
        argb: 'DDEBF7'
      }
    };

    records.forEach(v => {
      worksheet.addRow({
        facultyCode: v.facultyCode || '-',
        facultyName:
          v.facultyName || 'ไม่ระบุคณะ',
        studentId: v.studentId,
        name: v.name || 'ไม่ระบุชื่อ',
        activityName:
          v.activityName || 'ไม่ระบุกิจกรรม',
        hours: v.hours,
        date: v.date
          ? new Date(v.date).toLocaleString('th-TH')
          : '-',
        imageUrl:
          v.imageUrl || 'ไม่มีรูปภาพ'
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
    console.error(
      'Error exporting Excel:',
      error
    );

    res.status(500).send(
      'เกิดข้อผิดพลาดในการสร้างไฟล์ Excel'
    );
  }
});

// 5. หน้า Admin Dashboard - Modern Student Analytics UI
app.get('/admin', (req, res) => {
  res.send(`
<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Volunteer Admin • Student Analytics</title>

  <link
    href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"
    rel="stylesheet"
  >

  <link
    href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sarabun:wght@400;500;600;700&display=swap"
    rel="stylesheet"
  >

  <script src="https://unpkg.com/lucide@latest"></script>

  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>

  <style>
    :root{
      --bg:#f5f7fb;
      --surface:#fff;
      --surface-2:#f8fafc;
      --line:#e8edf5;
      --text:#172033;
      --muted:#718096;
      --primary:#635bff;
      --primary-2:#7c3aed;
      --success:#10b981;
      --warning:#f59e0b;
      --danger:#ef4444;
      --shadow:0 8px 30px rgba(15,23,42,.055);
      --radius:18px;
    }

    *{
      box-sizing:border-box
    }

    body{
      margin:0;
      background:
        radial-gradient(
          circle at 10% 0%,
          rgba(99,91,255,.08),
          transparent 28%
        ),
        var(--bg);
      color:var(--text);
      font-family:'Sarabun','Inter',sans-serif
    }

    .topbar{
      height:72px;
      background:rgba(255,255,255,.82);
      backdrop-filter:blur(18px);
      border-bottom:1px solid rgba(226,232,240,.8);
      position:sticky;
      top:0;
      z-index:20
    }

    .brand-mark{
      width:42px;
      height:42px;
      border-radius:13px;
      background:linear-gradient(
        135deg,
        var(--primary),
        var(--primary-2)
      );
      color:#fff;
      display:grid;
      place-items:center;
      box-shadow:
        0 8px 20px rgba(99,91,255,.22)
    }

    .page{
      max-width:1500px;
      margin:auto;
      padding:26px 24px 42px
    }

    .hero{
      display:flex;
      align-items:end;
      justify-content:space-between;
      gap:20px;
      margin-bottom:22px
    }

    .eyebrow{
      font-size:12px;
      font-weight:800;
      letter-spacing:.08em;
      text-transform:uppercase;
      color:var(--primary);
      margin-bottom:5px
    }

    .hero h1{
      font-size:clamp(25px,3vw,36px);
      font-weight:800;
      letter-spacing:-.03em;
      margin:0
    }

    .hero p{
      margin:5px 0 0;
      color:var(--muted);
      font-size:14px
    }

    .cardx{
      background:rgba(255,255,255,.92);
      border:1px solid var(--line);
      border-radius:var(--radius);
      box-shadow:var(--shadow)
    }

    .stat{
      padding:17px 18px;
      position:relative;
      overflow:hidden;
      min-height:110px
    }

    .stat:after{
      content:'';
      position:absolute;
      width:85px;
      height:85px;
      border-radius:50%;
      right:-30px;
      bottom:-38px;
      background:rgba(99,91,255,.06)
    }

    .stat-label{
      font-size:12px;
      color:var(--muted);
      font-weight:700
    }

    .stat-value{
      font-size:27px;
      font-weight:800;
      line-height:1.15;
      margin-top:7px
    }

    .stat-icon{
      width:38px;
      height:38px;
      border-radius:12px;
      display:grid;
      place-items:center;
      position:absolute;
      right:16px;
      top:16px
    }

    .icon-purple{
      background:#efedff;
      color:var(--primary)
    }

    .icon-green{
      background:#eafbf5;
      color:#059669
    }

    .icon-orange{
      background:#fff5df;
      color:#d97706
    }

    .icon-blue{
      background:#eaf4ff;
      color:#2563eb
    }

    .panel{
      padding:17px 18px
    }

    .panel-head{
      display:flex;
      align-items:center;
      justify-content:space-between;
      margin-bottom:10px
    }

    .panel-title{
      font-size:15px;
      font-weight:800
    }

    .panel-sub{
      font-size:12px;
      color:var(--muted)
    }

    .chart-box{
      height:190px;
      position:relative
    }

    .chart-box.donut{
      height:185px
    }

    .toolbar{
      padding:13px
    }

    .search-wrap{
      position:relative
    }

    .search-wrap svg{
      position:absolute;
      left:13px;
      top:11px;
      color:#94a3b8;
      width:17px
    }

    .search-wrap input{
      padding-left:40px;
      height:42px;
      border-radius:12px;
      border:1px solid var(--line);
      background:#fbfcfe
    }

    .form-select{
      height:42px;
      border-radius:12px;
      border-color:var(--line);
      background-color:#fbfcfe
    }

    .table-panel{
      overflow:hidden
    }

    .table thead th{
      font-size:11px;
      color:#8a94a6;
      text-transform:uppercase;
      letter-spacing:.06em;
      background:#fbfcfe;
      border-bottom:1px solid var(--line);
      padding:13px 14px;
      white-space:nowrap
    }

    .table tbody td{
      padding:12px 14px;
      border-bottom:1px solid #f0f3f8;
      vertical-align:middle
    }

    .table tbody tr{
      transition:.15s
    }

    .table tbody tr:hover{
      background:#fafbff
    }

    .student-btn{
      border:0;
      background:transparent;
      padding:0;
      text-align:left
    }

    .student-name{
      font-weight:800;
      color:var(--text);
      font-size:14px
    }

    .student-id{
      font-size:11px;
      color:var(--muted);
      margin-top:2px
    }

    .student-btn:hover .student-name{
      color:var(--primary)
    }

    .avatar{
      width:34px;
      height:34px;
      border-radius:11px;
      display:grid;
      place-items:center;
      background:linear-gradient(
        135deg,
        #eef2ff,
        #f5f3ff
      );
      color:var(--primary);
      font-weight:800;
      font-size:13px
    }

    .faculty-badge{
      display:inline-flex;
      align-items:center;
      gap:5px;
      padding:5px 9px;
      border-radius:9px;
      font-size:11px;
      font-weight:700;
      background:#f4f6fa;
      color:#556070;
      border:1px solid #e7ebf2
    }

    .hours-badge{
      display:inline-flex;
      padding:5px 9px;
      border-radius:999px;
      background:#eafbf5;
      color:#047857;
      font-weight:800;
      font-size:11px
    }

    .proof{
      width:38px;
      height:38px;
      object-fit:cover;
      border-radius:10px;
      border:1px solid var(--line);
      cursor:pointer
    }

    .no-proof{
      font-size:11px;
      color:#a0a8b6
    }

    .btn-smx{
      width:32px;
      height:32px;
      border-radius:9px;
      padding:0;
      display:inline-grid;
      place-items:center
    }

    .btn-primary-soft{
      background:#efedff;
      color:var(--primary);
      border:1px solid #ddd9ff
    }

    .btn-danger-soft{
      background:#fff0f0;
      color:#dc2626;
      border:1px solid #ffd7d7
    }

    .btn-view{
      background:#172033;
      color:#fff;
      border:0;
      border-radius:10px;
      padding:8px 12px;
      font-size:12px;
      font-weight:700
    }

    .empty{
      padding:42px 20px;
      text-align:center;
      color:var(--muted)
    }

    .modal-content{
      border:0;
      border-radius:22px;
      box-shadow:0 25px 70px rgba(15,23,42,.2);
      overflow:hidden
    }

    .modal-header{
      border-bottom:1px solid var(--line);
      padding:18px 20px
    }

    .modal-body{
      padding:20px
    }

    .profile-head{
      display:flex;
      gap:14px;
      align-items:center
    }

    .big-avatar{
      width:58px;
      height:58px;
      border-radius:18px;
      display:grid;
      place-items:center;
      background:linear-gradient(
        135deg,
        var(--primary),
        var(--primary-2)
      );
      color:white;
      font-size:20px;
      font-weight:800
    }

    .profile-name{
      font-size:20px;
      font-weight:800
    }

    .profile-meta{
      font-size:12px;
      color:var(--muted)
    }

    .mini-stat{
      padding:13px;
      border:1px solid var(--line);
      border-radius:14px;
      background:#fbfcfe
    }

    .mini-label{
      font-size:11px;
      color:var(--muted)
    }

    .mini-value{
      font-size:19px;
      font-weight:800;
      margin-top:3px
    }

    .student-table{
      font-size:12px
    }

    .student-table th{
      color:var(--muted);
      font-weight:700;
      background:#fafbfc
    }

    .student-table td,
    .student-table th{
      padding:10px;
      border-bottom:1px solid #edf1f6
    }

    .loading{
      opacity:.55;
      pointer-events:none
    }

    .toastx{
      position:fixed;
      right:20px;
      bottom:20px;
      z-index:100;
      display:none;
      padding:12px 16px;
      border-radius:12px;
      background:#172033;
      color:white;
      box-shadow:var(--shadow);
      font-size:13px
    }

    @media(max-width:767px){
      .page{
        padding:20px 14px
      }

      .hero{
        align-items:flex-start;
        flex-direction:column
      }

      .hero .actions{
        width:100%
      }

      .hero .actions a,
      .hero .actions button{
        flex:1
      }

      .chart-box{
        height:175px
      }

      .table thead th,
      .table tbody td{
        padding:10px
      }

      .hide-mobile{
        display:none!important
      }

      .topbar{
        height:64px
      }
    }
  </style>
</head>
<body>

  <header class="topbar d-flex align-items-center">
    <div class="container-fluid px-3 px-md-4 d-flex align-items-center justify-content-between">

      <div class="d-flex align-items-center gap-3">
        <div class="brand-mark">
          <i data-lucide="heart-handshake" width="21"></i>
        </div>

        <div>
          <div
            class="fw-bold"
            style="font-size:16px;line-height:1"
          >
            Volunteer Admin
          </div>

          <div
            style="font-size:11px;color:#8a94a6;margin-top:4px"
          >
            Student Analytics Dashboard
          </div>
        </div>
      </div>

      <div class="d-flex gap-2">

        <a
          href="/admin/export-excel"
          class="btn btn-sm btn-outline-success rounded-3 d-flex align-items-center gap-2 px-3"
        >
          <i data-lucide="download" width="15"></i>
          <span class="hide-mobile">Export</span>
        </a>

        <button
          class="btn btn-sm btn-dark rounded-3 d-flex align-items-center gap-2 px-3"
          onclick="loadData()"
        >
          <i data-lucide="refresh-cw" width="15"></i>
          <span class="hide-mobile">รีเฟรช</span>
        </button>

      </div>
    </div>
  </header>

  <main class="page">

    <section class="hero">

      <div>
        <div class="eyebrow">Overview</div>

        <h1>
          แดชบอร์ดชั่วโมงจิตอาสา
        </h1>

        <p>
          ดูภาพรวมและกดที่นักศึกษาเพื่อเปิดรายละเอียดรายบุคคล
        </p>
      </div>

      <div class="actions d-flex gap-2"></div>

    </section>

    <section class="row g-3 mb-3">

      <div class="col-6 col-xl-3">
        <div class="cardx stat">
          <div class="stat-label">
            รายการทั้งหมด
          </div>

          <div
            class="stat-value"
            id="statCount"
          >
            0
          </div>

          <div class="stat-icon icon-purple">
            <i data-lucide="clipboard-list" width="19"></i>
          </div>
        </div>
      </div>

      <div class="col-6 col-xl-3">
        <div class="cardx stat">

          <div class="stat-label">
            ชั่วโมงสะสม
          </div>

          <div
            class="stat-value"
            id="statHours"
          >
            0
          </div>

          <div class="stat-icon icon-green">
            <i data-lucide="clock-3" width="19"></i>
          </div>

        </div>
      </div>

      <div class="col-6 col-xl-3">
        <div class="cardx stat">

          <div class="stat-label">
            นักศึกษาที่เข้าร่วม
          </div>

          <div
            class="stat-value"
            id="statStudents"
          >
            0
          </div>

          <div class="stat-icon icon-blue">
            <i data-lucide="users" width="19"></i>
          </div>

        </div>
      </div>

      <div class="col-6 col-xl-3">
        <div class="cardx stat">

          <div class="stat-label">
            ชั่วโมงเฉลี่ย / รายการ
          </div>

          <div
            class="stat-value"
            id="statAverage"
          >
            0
          </div>

          <div class="stat-icon icon-orange">
            <i data-lucide="trending-up" width="19"></i>
          </div>

        </div>
      </div>

    </section>

    <section class="row g-3 mb-3">

      <div class="col-12 col-lg-7">

        <div class="cardx panel">

          <div class="panel-head">

            <div>
              <div class="panel-title">
                ชั่วโมงรวมแยกตามคณะ
              </div>

              <div class="panel-sub">
                เปรียบเทียบชั่วโมงสะสมของแต่ละคณะ
              </div>
            </div>

            <i
              data-lucide="bar-chart-3"
              width="18"
              style="color:var(--primary)"
            ></i>

          </div>

          <div class="chart-box">
            <canvas id="facultyChart"></canvas>
          </div>

        </div>

      </div>

      <div class="col-12 col-lg-5">

        <div class="cardx panel">

          <div class="panel-head">

            <div>
              <div class="panel-title">
                สัดส่วนการบันทึก
              </div>

              <div class="panel-sub">
                จำนวนรายการของแต่ละคณะ
              </div>
            </div>

            <i
              data-lucide="pie-chart"
              width="18"
              style="color:var(--primary)"
            ></i>

          </div>

          <div class="chart-box donut">
            <canvas id="ratioChart"></canvas>
          </div>

        </div>

      </div>

    </section>

    <section class="cardx table-panel">

      <div class="toolbar border-bottom">

        <div class="row g-2 align-items-center">

          <div class="col-12 col-lg-7">

            <div class="search-wrap">

              <i data-lucide="search"></i>

              <input
                id="searchInput"
                class="form-control"
                placeholder="ค้นหารหัสนักศึกษา ชื่อ คณะ หรือกิจกรรม..."
                oninput="filterTable()"
              >

            </div>

          </div>

          <div class="col-8 col-lg-3">

            <select
              id="facultyFilter"
              class="form-select"
              onchange="filterTable()"
            >
              <option value="">ทุกคณะ</option>
              <option value="01">01 • วิทยาศาสตร์ฯ</option>
              <option value="02">02 • บริหารธุรกิจฯ</option>
              <option value="03">03 • วิศวกรรมศาสตร์</option>
            </select>

          </div>

          <div class="col-4 col-lg-2 text-end">

            <span
              id="resultCount"
              class="panel-sub"
            ></span>

          </div>

        </div>

      </div>

      <div class="table-responsive">

        <table class="table mb-0 align-middle">

          <thead>
            <tr>
              <th>หลักฐาน</th>
              <th>นักศึกษา</th>
              <th>คณะ</th>
              <th>กิจกรรม</th>
              <th>ชั่วโมง</th>
              <th>วันที่</th>
              <th class="text-end">จัดการ</th>
            </tr>
          </thead>

          <tbody id="recordTableBody">

            <tr>
              <td
                colspan="7"
                class="empty"
              >
                กำลังโหลดข้อมูล...
              </td>
            </tr>

          </tbody>

        </table>

      </div>

    </section>

  </main>

  <div
    class="modal fade"
    id="studentModal"
    tabindex="-1"
  >

    <div
      class="modal-dialog modal-lg modal-dialog-centered modal-dialog-scrollable"
    >

      <div class="modal-content">

        <div class="modal-header">

          <div class="profile-head">

            <div
              class="big-avatar"
              id="studentAvatar"
            >
              ?
            </div>

            <div>

              <div
                class="profile-name"
                id="studentModalName"
              >
                ข้อมูลนักศึกษา
              </div>

              <div
                class="profile-meta"
                id="studentModalMeta"
              ></div>

            </div>

          </div>

          <button
            class="btn-close ms-auto"
            data-bs-dismiss="modal"
          ></button>

        </div>

        <div class="modal-body">

          <div class="row g-2 mb-3">

            <div class="col-4">

              <div class="mini-stat">

                <div class="mini-label">
                  ชั่วโมงสะสม
                </div>

                <div
                  class="mini-value text-success"
                  id="studentTotalHours"
                >
                  0
                </div>

              </div>

            </div>

            <div class="col-4">

              <div class="mini-stat">

                <div class="mini-label">
                  จำนวนรายการ
                </div>

                <div
                  class="mini-value"
                  id="studentRecordCount"
                >
                  0
                </div>

              </div>

            </div>

            <div class="col-4">

              <div class="mini-stat">

                <div class="mini-label">
                  ค่าเฉลี่ย / รายการ
                </div>

                <div
                  class="mini-value"
                  id="studentAverage"
                >
                  0
                </div>

              </div>

            </div>

          </div>

          <div class="table-responsive">

            <table class="w-100 student-table">

              <thead>
                <tr>
                  <th>วันที่</th>
                  <th>กิจกรรม</th>
                  <th>ชั่วโมง</th>
                  <th>หลักฐาน</th>
                </tr>
              </thead>

              <tbody id="studentRecordsBody"></tbody>

            </table>

          </div>

        </div>

      </div>

    </div>

  </div>

  <div
    class="modal fade"
    id="imageModal"
    tabindex="-1"
  >

    <div class="modal-dialog modal-dialog-centered">

      <div class="modal-content bg-transparent shadow-none">

        <div class="modal-body p-0 text-center">

          <img
            id="modalImg"
            src=""
            class="img-fluid rounded-4"
            style="max-height:80vh"
            alt="รูปหลักฐาน"
          >

        </div>

      </div>

    </div>

  </div>

  <div
    class="modal fade"
    id="editModal"
    tabindex="-1"
  >

    <div class="modal-dialog modal-dialog-centered">

      <div class="modal-content">

        <div class="modal-header">

          <h5 class="fw-bold mb-0">
            แก้ไขข้อมูลบันทึกจิตอาสา
          </h5>

          <button
            class="btn-close"
            data-bs-dismiss="modal"
          ></button>

        </div>

        <div class="modal-body">

          <form id="editForm">

            <input
              type="hidden"
              id="editId"
            >

            <div class="mb-3">

              <label class="form-label small fw-bold">
                คณะ
              </label>

              <select
                id="editFacultyCode"
                class="form-select"
                required
              >
                <option value="01">
                  01 - คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร
                </option>

                <option value="02">
                  02 - คณะบริหารธุรกิจและศิลปศาสตร์
                </option>

                <option value="03">
                  03 - คณะวิศวกรรมศาสตร์
                </option>
              </select>

            </div>

            <div class="mb-3">

              <label class="form-label small fw-bold">
                รหัสนักศึกษา
              </label>

              <input
                type="text"
                id="editStudentId"
                class="form-control"
                required
              >

            </div>

            <div class="mb-3">

              <label class="form-label small fw-bold">
                ชื่อ-นามสกุล
              </label>

              <input
                type="text"
                id="editName"
                class="form-control"
                required
              >

            </div>

            <div class="mb-3">

              <label class="form-label small fw-bold">
                จำนวนชั่วโมง
              </label>

              <input
                type="number"
                id="editHours"
                class="form-control"
                min="0.5"
                step="0.5"
                required
              >

            </div>

            <div class="mb-3">

              <label class="form-label small fw-bold">
                ชื่อกิจกรรม
              </label>

              <input
                type="text"
                id="editActivityName"
                class="form-control"
                required
              >

            </div>

            <div class="d-flex justify-content-end gap-2">

              <button
                type="button"
                class="btn btn-light"
                data-bs-dismiss="modal"
              >
                ยกเลิก
              </button>

              <button
                type="submit"
                class="btn btn-dark"
              >
                บันทึกการเปลี่ยนแปลง
              </button>

            </div>

          </form>

        </div>

      </div>

    </div>

  </div>

  <div
    class="toastx"
    id="toastx"
  ></div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>

  <script>
    let allRecords = [];
    let facultyChart = null;
    let ratioChart = null;

    const FACULTIES = {
      '01':'คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร',
      '02':'คณะบริหารธุรกิจและศิลปศาสตร์',
      '03':'คณะวิศวกรรมศาสตร์'
    };

    function esc(value){
      return String(value ?? '').replace(
        /[&<>'"]/g,
        c => ({
          '&':'&amp;',
          '<':'&lt;',
          '>':'&gt;',
          "'":'&#39;',
          '"':'&quot;'
        }[c])
      );
    }

    function initials(name){
      const t = (name || '?')
        .trim()
        .split(/\\s+/)
        .filter(Boolean);

      return t.length
        ? t.slice(0,2)
            .map(x => x[0])
            .join('')
            .toUpperCase()
        : '?';
    }

    function showToast(message){
      const el =
        document.getElementById('toastx');

      el.textContent = message;
      el.style.display = 'block';

      clearTimeout(window.__toast);

      window.__toast =
        setTimeout(
          () => el.style.display = 'none',
          2200
        );
    }

    async function loadData(){

      document.body.classList.add('loading');

      try{

        const res =
          await fetch('/api/admin/records');

        const result =
          await res.json();

        if(!result.success){
          throw new Error(
            result.error ||
            'โหลดข้อมูลไม่สำเร็จ'
          );
        }

        allRecords =
          result.data || [];

        updateStats(allRecords);
        renderCharts(allRecords);
        renderData(allRecords);

      }catch(err){

        console.error(err);

        alert(
          'ไม่สามารถดึงข้อมูลได้'
        );

      }finally{

        document.body.classList.remove(
          'loading'
        );

        lucide.createIcons();
      }
    }

    function updateStats(data){

      const totalHours =
        data.reduce(
          (s,r) =>
            s + (Number(r.hours) || 0),
          0
        );

      const students =
        new Set(
          data
            .map(r => r.studentId)
            .filter(Boolean)
        ).size;

      document.getElementById(
        'statCount'
      ).textContent =
        data.length.toLocaleString();

      document.getElementById(
        'statHours'
      ).textContent =
        totalHours.toLocaleString();

      document.getElementById(
        'statStudents'
      ).textContent =
        students.toLocaleString();

      document.getElementById(
        'statAverage'
      ).textContent =
        data.length
          ? (totalHours / data.length)
              .toFixed(1)
          : '0';
    }

    function renderCharts(data){

      const hours = {
        '01':0,
        '02':0,
        '03':0
      };

      const counts = {
        '01':0,
        '02':0,
        '03':0
      };

      data.forEach(r => {

        if(hours[r.facultyCode] !== undefined){

          hours[r.facultyCode] +=
            Number(r.hours) || 0;

          counts[r.facultyCode]++;
        }

      });

      if(facultyChart)
        facultyChart.destroy();

      if(ratioChart)
        ratioChart.destroy();

      facultyChart =
        new Chart(
          document.getElementById(
            'facultyChart'
          ),
          {
            type:'bar',

            data:{
              labels:[
                'วิทยาศาสตร์ฯ',
                'บริหารธุรกิจฯ',
                'วิศวกรรมศาสตร์'
              ],

              datasets:[
                {
                  data:[
                    hours['01'],
                    hours['02'],
                    hours['03']
                  ],
                  borderRadius:7,
                  barThickness:26
                }
              ]
            },

            options:{
              responsive:true,
              maintainAspectRatio:false,

              plugins:{
                legend:{
                  display:false
                },

                tooltip:{
                  displayColors:false,

                  callbacks:{
                    label:c =>
                      \` \${c.raw} ชั่วโมง\`
                  }
                }
              },

              scales:{
                x:{
                  grid:{
                    display:false
                  },

                  ticks:{
                    font:{
                      size:10
                    }
                  }
                },

                y:{
                  beginAtZero:true,

                  grid:{
                    color:'#eef1f6'
                  },

                  ticks:{
                    font:{
                      size:10
                    }
                  }
                }
              }
            }
          }
        );

      ratioChart =
        new Chart(
          document.getElementById(
            'ratioChart'
          ),
          {
            type:'doughnut',

            data:{
              labels:[
                'วิทยาศาสตร์ฯ',
                'บริหารธุรกิจฯ',
                'วิศวกรรมศาสตร์'
              ],

              datasets:[
                {
                  data:[
                    counts['01'],
                    counts['02'],
                    counts['03']
                  ],

                  borderWidth:0,
                  spacing:3
                }
              ]
            },

            options:{
              responsive:true,
              maintainAspectRatio:false,
              cutout:'70%',

              plugins:{
                legend:{
                  position:'bottom',

                  labels:{
                    boxWidth:9,
                    usePointStyle:true,
                    padding:12,

                    font:{
                      size:10
                    }
                  }
                }
              }
            }
          }
        );
    }
           function renderData(data) {
  const tbody = document.getElementById('recordTableBody');

  if (!data.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="empty">
          <i data-lucide="inbox" width="28"></i>
          <div class="mt-2">
            ไม่พบข้อมูลที่ตรงกับการค้นหา
          </div>
        </td>
      </tr>
    `;

    document.getElementById('resultCount').textContent = '0 รายการ';

    lucide.createIcons();
    return;
  }

  tbody.innerHTML = data.map(item => {
    const date = item.date
      ? new Date(item.date).toLocaleDateString('th-TH', {
          year: 'numeric',
          month: 'short',
          day: 'numeric'
        })
      : '-';

    const image = item.imageUrl
      ? `
        <img
          src="${esc(item.imageUrl)}"
          class="proof"
          onclick="openImage('${esc(item.imageUrl)}')"
          title="คลิกเพื่อดูรูป"
          alt="หลักฐาน"
        >
      `
      : `
        <span class="no-proof">
          ไม่มีรูป
        </span>
      `;

    const faculty =
      item.facultyName ||
      FACULTIES[item.facultyCode] ||
      'ไม่ระบุคณะ';

    return `
      <tr
        data-search="${esc([
          item.studentId,
          item.name,
          item.facultyName,
          item.facultyCode,
          item.activityName
        ].join(' '))}"
        data-faculty="${esc(item.facultyCode || '')}"
      >

        <td>
          ${image}
        </td>

        <td>
          <div class="d-flex align-items-center gap-2">

            <div class="avatar">
              ${esc(initials(item.name))}
            </div>

            <button
              class="student-btn"
              onclick="openStudent('${esc(item.studentId || '')}')"
            >
              <div class="student-name">
                ${esc(item.name || 'ไม่ระบุชื่อ')}
              </div>

              <div class="student-id">
                ${esc(item.studentId || '-')}
              </div>
            </button>

          </div>
        </td>

        <td>
          <span class="faculty-badge">
            ${esc(faculty)}
          </span>
        </td>

        <td>
          <div
            style="
              max-width:220px;
              font-size:13px;
              font-weight:600;
            "
          >
            ${esc(item.activityName || 'ไม่ระบุกิจกรรม')}
          </div>
        </td>

        <td>
          <span class="hours-badge">
            ${Number(item.hours) || 0} ชม.
          </span>
        </td>

        <td>
          <span
            style="
              font-size:12px;
              color:#687386;
            "
          >
            ${date}
          </span>
        </td>

        <td class="text-end">
          <div
            class="d-flex justify-content-end gap-1"
          >

            <button
              class="btn btn-smx btn-primary-soft"
              onclick="openEdit('${esc(item._id)}')"
              title="แก้ไข"
            >
              <i
                data-lucide="pencil"
                width="14"
              ></i>
            </button>

            <button
              class="btn btn-smx btn-danger-soft"
              onclick="deleteRecord('${esc(item._id)}')"
              title="ลบ"
            >
              <i
                data-lucide="trash-2"
                width="14"
              ></i>
            </button>

          </div>
        </td>

      </tr>
    `;
  }).join('');

  document.getElementById('resultCount').textContent =
    `${data.length.toLocaleString()} รายการ`;

  lucide.createIcons();
}

      const search =
        document.getElementById(
          'searchInput'
        ).value
          .trim()
          .toLowerCase();

      const faculty =
        document.getElementById(
          'facultyFilter'
        ).value;

      const filtered =
        allRecords.filter(item => {

          const text = [
            item.studentId,
            item.name,
            item.facultyName,
            item.facultyCode,
            item.activityName
          ]
            .join(' ')
            .toLowerCase();

          const matchSearch =
            !search ||
            text.includes(search);

          const matchFaculty =
            !faculty ||
            item.facultyCode === faculty;

          return (
            matchSearch &&
            matchFaculty
          );
        });

      renderData(filtered);
    }

    async function openStudent(studentId){

      if(!studentId){
        showToast(
          'ไม่พบรหัสนักศึกษา'
        );
        return;
      }

      try{

        const res =
          await fetch(
            `/api/admin/students/${encodeURIComponent(
              studentId
            )}`
          );

        const result =
          await res.json();

        if(!result.success){
          throw new Error(
            result.message ||
            'ไม่พบข้อมูลนักศึกษา'
          );
        }

        const student =
          result.data;

        document.getElementById(
          'studentAvatar'
        ).textContent =
          initials(student.name);

        document.getElementById(
          'studentModalName'
        ).textContent =
          student.name ||
          'ไม่ระบุชื่อ';

        document.getElementById(
          'studentModalMeta'
        ).textContent =
          `${student.studentId || '-'} • ${
            student.facultyName ||
            'ไม่ระบุคณะ'
          }`;

        document.getElementById(
          'studentTotalHours'
        ).textContent =
          `${student.totalHours || 0} ชม.`;

        document.getElementById(
          'studentRecordCount'
        ).textContent =
          student.totalRecords || 0;

        document.getElementById(
          'studentAverage'
        ).textContent =
          student.totalRecords
            ? (
                student.totalHours /
                student.totalRecords
              ).toFixed(1)
            : '0';

        const tbody =
          document.getElementById(
            'studentRecordsBody'
          );

        tbody.innerHTML =
          (student.records || [])
            .map(record => {

              const date =
                record.date
                  ? new Date(
                      record.date
                    ).toLocaleDateString(
                      'th-TH',
                      {
                        year:'numeric',
                        month:'short',
                        day:'numeric'
                      }
                    )
                  : '-';

              const proof =
                record.imageUrl
                  ? `
                    <button
                      class="btn btn-smx btn-primary-soft"
                      onclick="openImage('${esc(
                        record.imageUrl
                      )}')"
                      title="ดูหลักฐาน"
                    >
                      <i
                        data-lucide="image"
                        width="14"
                      ></i>
                    </button>
                  `
                  : `
                    <span
                      class="no-proof"
                    >
                      ไม่มีรูป
                    </span>
                  `;

              return `
                <tr>

                  <td>
                    ${date}
                  </td>

                  <td>
                    <div
                      style="
                        font-weight:600;
                        color:#273246;
                      "
                    >
                      ${esc(
                        record.activityName ||
                        'ไม่ระบุกิจกรรม'
                      )}
                    </div>
                  </td>

                  <td>
                    <span class="hours-badge">
                      ${Number(record.hours) || 0}
                      ชม.
                    </span>
                  </td>

                  <td>
                    ${proof}
                  </td>

                </tr>
              `;

            })
            .join('');

        if(!student.records.length){

          tbody.innerHTML = `
            <tr>
              <td
                colspan="4"
                class="text-center text-muted py-4"
              >
                ไม่พบประวัติกิจกรรม
              </td>
            </tr>
          `;

        }

        lucide.createIcons();

        bootstrap.Modal
          .getOrCreateInstance(
            document.getElementById(
              'studentModal'
            )
          )
          .show();

      }catch(err){

        console.error(err);

        showToast(
          err.message ||
          'ไม่สามารถโหลดข้อมูลนักศึกษาได้'
        );
      }
    }

    function openImage(url){

      if(!url) return;

      document.getElementById(
        'modalImg'
      ).src = url;

      bootstrap.Modal
        .getOrCreateInstance(
          document.getElementById(
            'imageModal'
          )
        )
        .show();
    }

    function openEdit(id){

      const item =
        allRecords.find(
          r => r._id === id
        );

      if(!item){
        showToast(
          'ไม่พบข้อมูลรายการ'
        );
        return;
      }

      document.getElementById(
        'editId'
      ).value = item._id;

      document.getElementById(
        'editFacultyCode'
      ).value =
        item.facultyCode || '01';

      document.getElementById(
        'editStudentId'
      ).value =
        item.studentId || '';

      document.getElementById(
        'editName'
      ).value =
        item.name || '';

      document.getElementById(
        'editHours'
      ).value =
        item.hours || 0;

      document.getElementById(
        'editActivityName'
      ).value =
        item.activityName || '';

      bootstrap.Modal
        .getOrCreateInstance(
          document.getElementById(
            'editModal'
          )
        )
        .show();
    }

    document.getElementById(
      'editForm'
    ).addEventListener(
      'submit',
      async function(e){

        e.preventDefault();

        const id =
          document.getElementById(
            'editId'
          ).value;

        const payload = {
          facultyCode:
            document.getElementById(
              'editFacultyCode'
            ).value,

          studentId:
            document.getElementById(
              'editStudentId'
            ).value.trim(),

          name:
            document.getElementById(
              'editName'
            ).value.trim(),

          hours:
            Number(
              document.getElementById(
                'editHours'
              ).value
            ),

          activityName:
            document.getElementById(
              'editActivityName'
            ).value.trim()
        };

        try{

          const res =
            await fetch(
              `/api/admin/records/${encodeURIComponent(
                id
              )}`,
              {
                method:'PUT',

                headers:{
                  'Content-Type':
                    'application/json'
                },

                body:
                  JSON.stringify(payload)
              }
            );

          const result =
            await res.json();

          if(!result.success){

            throw new Error(
              result.error ||
              result.message ||
              'บันทึกไม่สำเร็จ'
            );
          }

          bootstrap.Modal
            .getOrCreateInstance(
              document.getElementById(
                'editModal'
              )
            )
            .hide();

          showToast(
            'บันทึกข้อมูลเรียบร้อยแล้ว'
          );

          await loadData();

        }catch(err){

          console.error(err);

          showToast(
            err.message ||
            'เกิดข้อผิดพลาดในการบันทึก'
          );
        }
      }
    );

    async function deleteRecord(id){

      const item =
        allRecords.find(
          r => r._id === id
        );

      if(!item){
        showToast(
          'ไม่พบข้อมูลรายการ'
        );
        return;
      }

      const ok =
        confirm(
          `ต้องการลบรายการของ ${
            item.name || 'นักศึกษา'
          } ใช่หรือไม่?\\n\\nกิจกรรม: ${
            item.activityName ||
            'ไม่ระบุกิจกรรม'
          }\\nชั่วโมง: ${
            item.hours || 0
          } ชั่วโมง`
        );

      if(!ok) return;

      try{

        const res =
          await fetch(
            `/api/admin/records/${encodeURIComponent(
              id
            )}`,
            {
              method:'DELETE'
            }
          );

        const result =
          await res.json();

        if(!result.success){

          throw new Error(
            result.error ||
            result.message ||
            'ลบข้อมูลไม่สำเร็จ'
          );
        }

        showToast(
          'ลบข้อมูลเรียบร้อยแล้ว'
        );

        await loadData();

      }catch(err){

        console.error(err);

        showToast(
          err.message ||
          'เกิดข้อผิดพลาดในการลบข้อมูล'
        );
      }
    }

    document.addEventListener(
      'DOMContentLoaded',
      () => {

        lucide.createIcons();

        loadData();

      }
    );
  </script>

</body>
</html>
  `);
});

// ==========================================
// 💬 LINE BOT HANDLER
// ==========================================
// 💬 LINE BOT HANDLER
// ==========================================

async function replyTextMsg(replyToken, text) {
  try {
    if (line.messagingApi) {
      return await client.replyMessage({
        replyToken,
        messages: [
          {
            type: 'text',
            text
          }
        ]
      });
    }

    return await client.replyMessage(replyToken, {
      type: 'text',
      text
    });

  } catch (error) {
    console.error('Error replying message:', error);
  }
}

async function handleImageMessage(event) {
  const userId = event.source.userId;

  try {
    let imageBuffer;

    if (line.messagingApi) {
      const stream =
        await blobClient.getMessageContent(
          event.message.id
        );

      const chunks = [];

      for await (const chunk of stream) {
        chunks.push(chunk);
      }

      imageBuffer =
        Buffer.concat(chunks);

    } else {
      const stream =
        await client.getMessageContent(
          event.message.id
        );

      const chunks = [];

      for await (const chunk of stream) {
        chunks.push(chunk);
      }

      imageBuffer =
        Buffer.concat(chunks);
    }

    const uploadResult =
      await new Promise(
        (resolve, reject) => {

          const uploadStream =
            cloudinary.uploader.upload_stream(
              {
                folder:
                  'volunteer_proofs'
              },

              (error, result) => {

                if (error) {
                  reject(error);
                } else {
                  resolve(result);
                }

              }
            );

          uploadStream.end(
            imageBuffer
          );
        }
      );

    const imageUrl =
      uploadResult.secure_url;

    await TempImage.findOneAndUpdate(
      { userId },
      {
        imageUrl,
        createdAt: new Date()
      },
      {
        upsert: true,
        new: true
      }
    );

    return replyTextMsg(
      event.replyToken,
      '📷 ได้รับรูปภาพหลักฐานเรียบร้อยแล้วครับ!\n\nกรุณาพิมพ์บันทึกชั่วโมงต่อได้เลย เช่น:\nบันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด'
    );

  } catch (error) {

    console.error(
      'Error handling image with Cloudinary:',
      error
    );

    return replyTextMsg(
      event.replyToken,
      '❌ ไม่สามารถบันทึกรูปภาพได้ กรุณาลองส่งใหม่อีกครั้ง'
    );
  }
}

async function handleEvent(event) {

  if (
    event.type === 'message' &&
    event.message.type === 'image'
  ) {
    return handleImageMessage(event);
  }

  if (
    event.type !== 'message' ||
    event.message.type !== 'text'
  ) {
    return Promise.resolve(null);
  }

  const userText =
    event.message.text.trim();

  const userId =
    event.source.userId;

  // ==========================================
  // 1. คำสั่ง "เช็คชั่วโมง"
  // ==========================================

  if (
    userText.startsWith(
      'เช็คชั่วโมง'
    )
  ) {

    const parts =
      userText
        .split(/\s+/)
        .filter(
          p => p.trim() !== ''
        );

    const studentId =
      parts[1];

    if (!studentId) {

      return replyTextMsg(
        event.replyToken,
        '❌ กรุณาระบุรหัสนักศึกษา เช่น:\nเช็คชั่วโมง 6501234567'
      );
    }

    try {

      const records =
        await Volunteer.find({
          studentId: studentId
        });

      if (records.length === 0) {

        return replyTextMsg(
          event.replyToken,
          `🔍 ไม่พบข้อมูลการบันทึกชั่วโมงของรหัส: ${studentId}`
        );
      }

      const totalHours =
        records.reduce(
          (sum, item) =>
            sum +
            (Number(item.hours) || 0),
          0
        );

      const validRecord =
        [...records]
          .reverse()
          .find(
            r =>
              r.name &&
              r.name !==
                'ไม่ระบุชื่อ'
          );

      const studentName =
        validRecord
          ? validRecord.name
          : (
              records[0].name ||
              'ไม่ระบุชื่อ'
            );

      const facultyName =
        validRecord
          ? (
              validRecord.facultyName ||
              'ไม่ระบุ'
            )
          : 'ไม่ระบุ';

      const replyText =
        `📊 สรุปชั่วโมงจิตอาสา\n\n` +
        `👤 ชื่อ: ${studentName}\n` +
        `🆔 รหัส: ${studentId}\n` +
        `🏛 คณะ: ${facultyName}\n` +
        `📝 บันทึกทั้งหมด: ${records.length} ครั้ง\n` +
        `⏱ ชั่วโมงสะสมรวม: ${totalHours} ชั่วโมง`;

      return replyTextMsg(
        event.replyToken,
        replyText
      );

    } catch (error) {

      console.error(
        'Error fetching data:',
        error
      );

      return replyTextMsg(
        event.replyToken,
        '❌ เกิดข้อผิดพลาดในการดึงข้อมูล กรุณาลองใหม่อีกครั้ง'
      );
    }
  }

  // ==========================================
  // 2. คำสั่ง "บันทึก"
  // ==========================================

  if (
    userText.startsWith(
      'บันทึก'
    )
  ) {

    const parts =
      userText
        .split(/\s+/)
        .filter(
          p => p.trim() !== ''
        );

    if (parts.length < 6) {

      return replyTextMsg(
        event.replyToken,
        '❌ รูปแบบคำสั่งไม่ถูกต้อง!\nกรุณาพิมพ์: บันทึก <รหัสคณะ> <รหัสนักศึกษา> <ชื่อ-นามสกุล> <จำนวนชั่วโมง> <ชื่อกิจกรรม>\n\n' +
        '🏛 รหัสคณะ:\n' +
        '01 = วิทยาศาสตร์และเทคโนโลยีการเกษตร\n' +
        '02 = บริหารธุรกิจและศิลปศาสตร์\n' +
        '03 = วิศวกรรมศาสตร์\n\n' +
        'ตัวอย่าง:\n' +
        'บันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด'
      );
    }

    const facultyCode =
      parts[1];

    const studentId =
      parts[2];

    let hoursIndex = -1;

    for (
      let i = 3;
      i < parts.length - 1;
      i++
    ) {

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

    const name =
      parts
        .slice(
          3,
          hoursIndex
        )
        .join(' ')
        .trim();

    const hours =
      parseFloat(
        parts[hoursIndex]
      );

    const activityName =
      parts
        .slice(
          hoursIndex + 1
        )
        .join(' ')
        .trim();

    if (!FACULTY_MAP[facultyCode]) {

      return replyTextMsg(
        event.replyToken,
        '❌ รหัสคณะไม่ถูกต้อง!\nกรุณาใช้รหัสคณะดังนี้:\n01 = คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร\n02 = คณะบริหารธุรกิจและศิลปศาสตร์\n03 = คณะวิศวกรรมศาสตร์'
      );
    }

    const facultyName =
      FACULTY_MAP[facultyCode];

    if (!name) {

      return replyTextMsg(
        event.replyToken,
        '❌ ไม่พบชื่อ-นามสกุล กรุณาตรวจสอบรูปแบบอีกครั้ง'
      );
    }

    if (
      isNaN(hours) ||
      hours <= 0
    ) {

      return replyTextMsg(
        event.replyToken,
        '❌ จำนวนชั่วโมงต้องเป็นตัวเลขที่มากกว่า 0'
      );
    }

    if (!activityName) {

      return replyTextMsg(
        event.replyToken,
        '❌ กรุณาระบุชื่อกิจกรรมต่อท้ายด้วยครับ'
      );
    }

    try {

      const tempImg =
        await TempImage.findOneAndDelete({
          userId
        });

      const imageUrl =
        tempImg
          ? tempImg.imageUrl
          : '';

      const newRecord =
        new Volunteer({
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

      const allRecords =
        await Volunteer.find({
          studentId
        });

      const totalHours =
        allRecords.reduce(
          (sum, item) =>
            sum +
            (Number(item.hours) || 0),
          0
        );

      let replyText =
        `✅ บันทึกชั่วโมงจิตอาสาสำเร็จ!\n\n` +
        `👤 ชื่อ: ${name}\n` +
        `🆔 รหัส: ${studentId}\n` +
        `🏛 คณะ: ${facultyName}\n` +
        `📌 กิจกรรม: ${activityName}\n` +
        `⏱ บันทึกเพิ่ม: ${hours} ชั่วโมง\n` +
        `📊 ชั่วโมงสะสมรวม: ${totalHours} ชั่วโมง`;

      if (imageUrl) {

        replyText +=
          `\n📷 แนบรูปภาพหลักฐานเรียบร้อยแล้ว`;

      } else {

        replyText +=
          `\n⚠️ (บันทึกโดยไม่มีรูปภาพหลักฐาน)`;
      }

      return replyTextMsg(
        event.replyToken,
        replyText
      );

    } catch (error) {

      console.error(
        'Error saving to DB:',
        error
      );

      return replyTextMsg(
        event.replyToken,
        '❌ เกิดข้อผิดพลาดในการบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง'
      );
    }
  }

  // ==========================================
  // 3. ข้อความแนะนำการใช้งาน
  // ==========================================

  const helpText =
    `👋 ยินดีต้อนรับสู่ระบบบันทึกชั่วโมงจิตอาสา\n\n` +
    `📌 ขั้นตอนการใช้งาน:\n` +
    `1️⃣ (ถ้ามี) ส่งรูปภาพหลักฐานการทำกิจกรรม\n` +
    `2️⃣ พิมพ์บันทึกชั่วโมงตามรูปแบบ:\n` +
    `บันทึก <รหัสคณะ> <รหัสประจำตัว> <ชื่อ-นามสกุล> <ชั่วโมง> <ชื่อกิจกรรม>\n\n` +
    `🏛 รหัสคณะ:\n` +
    `01 = คณะวิทยาศาสตร์และเทคโนโลยีการเกษตร\n` +
    `02 = คณะบริหารธุรกิจและศิลปศาสตร์\n` +
    `03 = คณะวิศวกรรมศาสตร์\n\n` +
    `💡 ตัวอย่าง:\n` +
    `บันทึก 01 6501234567 สมชาย ใจดี 4 ทำความสะอาดวัด\n\n` +
    `🔍 เช็คชั่วโมงสะสม:\n` +
    `พิมพ์: เช็คชั่วโมง <รหัสนักศึกษา>`;

  return replyTextMsg(
    event.replyToken,
    helpText
  );
}

const PORT =
  process.env.PORT || 3000;

app.listen(
  PORT,
  () => {
    console.log(
      `🚀 Server เปิดทำงานแล้วที่ Port ${PORT}`
    );
  }
);