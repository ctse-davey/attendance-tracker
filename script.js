const startDate = new Date("2025-04-21");
const weeks = 49;
const students = {
    "第1組 - 碧蓮 ★ 文信": ["碧蓮", "文信", "焦華彬", "黃潤萍", "胡灝文", "羅青玲", "廖喜兒", "任麥寶燕", "羅美好", "盧寶萍BoBo"],
    "第2組 - 桂蘭 ★ Jenny": ["桂蘭", "Jenny", "蕭慧明", "周貴珍", "林寶珍", "吳建容", "梁漪好", "黎桂英", "方娟梅"],
    "第3組 - 彭姑娘 ★ Winnie": ["彭姑娘", "Winnie", "何天嬌", "溫房嬌", "朱寶兒", "李玉華", "符志蓮Lilian", "蔡麗萍Doris", "鄭瑞琼Gloria", "黃綺華"]
};

const monthMap = {};
for (let i = 0; i < weeks; i++) {
    const date = new Date(startDate);
    date.setDate(date.getDate() + i * 7);
    const label = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!monthMap[label]) {
        monthMap[label] = [];
    }
    monthMap[label].push({ week: i + 1, date: date.toISOString().split("T")[0] });
}

const monthTabs = document.getElementById("monthTabs");
const tablesContainer = document.getElementById("tablesContainer");

// Export Filter Controls
const exportControls = document.createElement("div");
exportControls.style.margin = "20px";
exportControls.innerHTML = `
<label>匯出類型：</label>
<select id="exportType">
  <option value="all">全部</option>
  <option value="group">群組名稱</option>
  <option value="month">月份 (2025-05)</option>
  <option value="week">週次 (W12)</option>
</select>
<select id="exportKeyword" style="margin-left:10px;"></select>
<button id="exportExcel" style="margin-left:10px;">匯出為 Excel</button>
`;
document.body.insertBefore(exportControls, document.getElementById("tablesContainer"));
// =========================
// Sync to Google Sheet via Apps Script Web App
// =========================
const syncButton = document.createElement("button");
syncButton.textContent = "同步至 Google 試算表";
syncButton.style.marginLeft = "10px";
exportControls.appendChild(syncButton);

syncButton.onclick = async () => {
    const rows = [["群組","學生","週次","日期","出席"]];
    Object.entries(students).forEach(([groupName, studentList]) => {
        studentList.forEach(student => {
            for (let i = 1; i <= weeks; i++) {
                const checkbox = document.querySelector(`input[name='${student}-w${i}']`);
                if (!checkbox) continue;
                const date = new Date(startDate);
                date.setDate(date.getDate() + (i - 1) * 7);
                const ymd = date.toISOString().split("T")[0];
                const weekStr = "W" + i;
                rows.push([groupName, student, weekStr, ymd, checkbox.checked ? "1" : "0"]);
            }
        });
    });
    try {
        const resp = await fetch("https://script.google.com/macros/s/AKfycbyLKxAMq7_w5VKbLfbqspwZ2JdKuqagql7xY3y_JpdEF1R8CvQWV0IepFCPCjrx9wJx/exec", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(rows)
        });
        const result = await resp.json();
        if (resp.ok && result.status === 'success') {
            alert("同步成功！");
        } else {
            alert("同步失敗: " + (result.message || resp.statusText));
        }
    } catch (err) {
        alert("同步錯誤: " + err.message);
    }
};


// =========================
// Populate exportKeyword options based on exportType
// =========================
// ... rest of original file ...
