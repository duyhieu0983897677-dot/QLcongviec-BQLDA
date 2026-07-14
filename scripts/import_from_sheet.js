// Đọc file CSV export từ Google Sheet "Bảng tổng hợp kế hoạch và tiến độ"
// và chuyển thành seed_data.json đúng schema tasks[]/users[] cho app mới.
// Chạy 1 lần: node scripts/import_from_sheet.js
const fs = require('fs');
const path = require('path');

const CSV_PATH = path.join(__dirname, '..', 'data_source', 'ke_hoach_tien_do_export.csv');
const OUT_PATH = path.join(__dirname, '..', 'data_source', 'seed_data.json');
const REPORT_PATH = path.join(__dirname, '..', 'data_source', 'import_report.txt');

function parseCSV(text) {
    const rows = [];
    let row = [], field = '', inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (inQuotes) {
            if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; } }
            else field += c;
        } else {
            if (c === '"') inQuotes = true;
            else if (c === ',') { row.push(field); field = ''; }
            else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
            else if (c === '\r') { /* ignore */ }
            else field += c;
        }
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows;
}

const MONTH_MAP = { 'thg1': 1, 'thg2': 2, 'thg3': 3, 'thg4': 4, 'thg5': 5, 'thg6': 6, 'thg7': 7, 'thg8': 8, 'thg9': 9, 'thg10': 10, 'thg11': 11, 'thg12': 12 };

// Trả về {iso, ok} - parse linh hoạt nhiều định dạng ngày lẫn lộn trong sheet gốc.
// Năm mặc định 2026 nếu không có năm trong chuỗi (theo bối cảnh dự án).
function parseFlexDate(raw) {
    const s = String(raw || '').trim();
    if (!s) return { iso: '', ok: true };

    // dd-ThgN hoặc dd-ThgN-yyyy (VD: 04-Thg3)
    let m = s.toLowerCase().match(/^(\d{1,2})-(thg\d{1,2})(?:-(\d{4}))?$/);
    if (m) {
        const mo = MONTH_MAP[m[2]];
        if (mo) return { iso: `${m[3] || 2026}-${String(mo).padStart(2, '0')}-${m[1].padStart(2, '0')}`, ok: true };
    }

    // dd-mm-yyyy hoặc d-m-yyyy (số thuần, gạch ngang) — khác với dạng dd-ThgN ở trên
    m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
    if (m) {
        const year = m[3].length === 2 ? '20' + m[3] : m[3];
        return { iso: `${year}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`, ok: true };
    }

    // dd/mm/yyyy hoặc dd/mm/yy hoặc dd/mm
    m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
    if (m) {
        let year = m[3] ? (m[3].length === 2 ? '20' + m[3] : m[3]) : '2026';
        return { iso: `${year}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`, ok: true };
    }

    return { iso: '', ok: false };
}

const csvText = fs.readFileSync(CSV_PATH, 'utf8');
const rows = parseCSV(csvText);

const COL = { supervisor: 0, code: 1, name: 2, target: 3, duration: 4, start: 5, end: 6, percent: 38, note: 39, comment: 40 };

let taskId = 1;
const tasks = [];
const supervisorCounts = {}; // tên gốc -> số lần xuất hiện
const unparsedDates = []; // {row, raw}
let currentSupervisor = '';
let currentGroupId = null;

for (let i = 9; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length < 3) continue;
    const supervisorRaw = (r[COL.supervisor] || '').trim();
    const name = (r[COL.name] || '').trim();
    const target = (r[COL.target] || '').trim();
    const duration = (r[COL.duration] || '').trim();
    const startRaw = (r[COL.start] || '').trim();
    const endRaw = (r[COL.end] || '').trim();
    const percentRaw = (r[COL.percent] || '').trim();
    const note = (r[COL.note] || '').trim();

    if (!supervisorRaw && !name && !duration && !startRaw && !endRaw && !percentRaw && !note) continue; // dòng trắng hoàn toàn
    if (!name) continue; // không có tên thì bỏ qua (dòng lỗi)

    if (supervisorRaw) {
        currentSupervisor = supervisorRaw.split('\n')[0].trim(); // dòng đầu của ô (bỏ email nếu xuống dòng)
        supervisorCounts[currentSupervisor] = (supervisorCounts[currentSupervisor] || 0) + 1;
    }

    const hasScheduleData = !!(duration || startRaw || endRaw || percentRaw || note);

    if (!hasScheduleData) {
        // Dòng Nhóm (header) — không có dữ liệu tiến độ, chỉ là tên phân đoạn/hạng mục
        currentGroupId = taskId;
        tasks.push({
            id: taskId++,
            isHeader: true,
            name,
            groupId: null,
            supervisorName: '',
            target: '', duration: '', startDate: '', endDate: '', dateRawStart: '', dateRawEnd: '',
            percent: 0, note: '', dailyLogs: []
        });
        continue;
    }

    // Dòng Việc (task lá)
    const startParsed = parseFlexDate(startRaw);
    const endParsed = parseFlexDate(endRaw);
    if (startRaw && !startParsed.ok) unparsedDates.push({ row: i + 1, field: 'Bắt đầu', raw: startRaw, task: name });
    if (endRaw && !endParsed.ok) unparsedDates.push({ row: i + 1, field: 'Kết thúc', raw: endRaw, task: name });

    let percent = 0;
    const pctMatch = percentRaw.match(/(\d+(?:[.,]\d+)?)\s*%/);
    if (pctMatch) percent = parseFloat(pctMatch[1].replace(',', '.'));

    tasks.push({
        id: taskId++,
        isHeader: false,
        name,
        groupId: currentGroupId,
        supervisorName: currentSupervisor,
        target,
        duration,
        startDate: startParsed.iso,
        endDate: endParsed.iso,
        dateRawStart: startParsed.ok ? '' : startRaw,
        dateRawEnd: endParsed.ok ? '' : endRaw,
        percent,
        note,
        dailyLogs: []
    });
}

// Danh sách Giám sát duy nhất (sắp theo số lần xuất hiện giảm dần) để người dùng soát tên trùng
const supervisorList = Object.entries(supervisorCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([name, count]) => ({ name, count }));

fs.writeFileSync(OUT_PATH, JSON.stringify({ tasks }, null, 2), 'utf8');

const headerCount = tasks.filter(t => t.isHeader).length;
const taskCount = tasks.length - headerCount;
const report = [
    `Tổng số dòng đã import: ${tasks.length}`,
    `  - Dòng Nhóm (header): ${headerCount}`,
    `  - Dòng Việc (task): ${taskCount}`,
    ``,
    `Số ngày không parse được (giữ nguyên text gốc): ${unparsedDates.length}`,
    ...unparsedDates.slice(0, 30).map(d => `  - Dòng CSV ${d.row} [${d.field}] "${d.raw}" (việc: ${d.task.slice(0, 50)})`),
    unparsedDates.length > 30 ? `  ... và ${unparsedDates.length - 30} dòng khác` : '',
    ``,
    `Danh sách tên Giám sát xuất hiện trong sheet (${supervisorList.length} biến thể) — cần soát gộp tên trùng:`,
    ...supervisorList.map(s => `  - "${s.name}" (${s.count} việc)`)
].join('\n');

fs.writeFileSync(REPORT_PATH, report, 'utf8');
console.log(report);
console.log(`\nĐã ghi: ${OUT_PATH}`);
console.log(`Báo cáo: ${REPORT_PATH}`);
