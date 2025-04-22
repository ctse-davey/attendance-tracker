
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
// Populate exportKeyword options based on exportType
// =========================
const exportTypeSelect = document.getElementById("exportType");
const exportKeywordSelect = document.getElementById("exportKeyword");

function updateKeywordOptions() {
    const type = exportTypeSelect.value;
    exportKeywordSelect.innerHTML = "";
    if (type === "group") {
        Object.keys(students).forEach(groupName => {
            const opt = document.createElement("option");
            opt.value = groupName;
            opt.text = groupName;
            exportKeywordSelect.appendChild(opt);
        });
        exportKeywordSelect.style.display = "inline-block";
    } else if (type === "month") {
        Object.keys(monthMap).forEach(monthKey => {
            const opt = document.createElement("option");
            opt.value = monthKey;
            opt.text = monthKey;
            exportKeywordSelect.appendChild(opt);
        });
        exportKeywordSelect.style.display = "inline-block";
    } else if (type === "week") {
        for (let i = 1; i <= weeks; i++) {
            const w = "W" + i;
            const opt = document.createElement("option");
            opt.value = w;
            opt.text = w;
            exportKeywordSelect.appendChild(opt);
        }
        exportKeywordSelect.style.display = "inline-block";
    } else {
        exportKeywordSelect.style.display = "none";
    }
}

// Initialize options on load
exportTypeSelect.addEventListener("change", updateKeywordOptions);
updateKeywordOptions();


document.getElementById("exportExcel").onclick = () => {
    const type = document.getElementById("exportType").value;
    const keyword = document.getElementById("exportKeyword").value.trim();
    const data = [["群組", "學生", "週次", "日期", "出席"]];

    Object.entries(students).forEach(([groupName, studentList]) => {
        if (type === "group" && !groupName.includes(keyword)) return;

        studentList.forEach(student => {
            for (let i = 1; i <= weeks; i++) {
                const checkbox = document.querySelector(`input[name='${student}-w${i}']`);
                if (!checkbox) continue;

                const date = new Date(startDate);
                date.setDate(date.getDate() + (i - 1) * 7);
                const ymd = date.toISOString().split("T")[0];
                const ym = ymd.slice(0, 7);
                const weekStr = "W" + i;

                if (type === "month" && !ym.includes(keyword)) continue;
                if (type === "week" && weekStr !== keyword) continue;

                data.push([groupName, student, weekStr, ymd, checkbox.checked ? "1" : "0"]);
            }
        });
    });

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "出席資料");
    XLSX.writeFile(workbook, "attendance_filtered_export.xlsx");
};

// Create tables
Object.keys(monthMap).forEach((monthKey, idx) => {
    const tab = document.createElement("div");
    tab.className = "tab" + (idx === 0 ? " active" : "");
    tab.innerText = monthKey;
    tab.dataset.target = monthKey;
    tab.onclick = () => {
        document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
        document.querySelectorAll(".month-table").forEach(d => d.style.display = "none");
        tab.classList.add("active");
        document.getElementById(monthKey).style.display = "block";
    };
    monthTabs.appendChild(tab);

    const container = document.createElement("div");
    container.className = "month-table";
    container.id = monthKey;
    container.style.display = idx === 0 ? "block" : "none";

    for (const [groupName, groupStudents] of Object.entries(students)) {
        const group = document.createElement("div");
        group.className = "group";
        group.innerHTML = `<h2>${groupName}</h2>`;

        const table = document.createElement("table");
        const thead = document.createElement("thead");
        const headRow = document.createElement("tr");
        headRow.innerHTML = "<th>姓名</th>" + monthMap[monthKey].map(w => {
            const d = new Date(w.date);
            return `<th>W${w.week} - ${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日</th>`;
        }).join("");
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        groupStudents.forEach(name => {
            const row = document.createElement("tr");
            row.innerHTML = `<td>${name}</td>` + monthMap[monthKey].map(w => {
                return `<td><input type="checkbox" name="${name}-w${w.week}"></td>`;
            }).join("");
            tbody.appendChild(row);
        });

        const countRow = document.createElement("tr");
        countRow.innerHTML = `<td><strong>出席人數</strong></td>` + monthMap[monthKey].map(w => {
            const id = `count-${groupName.replace(/\s+/g, '')}-w${w.week}`;
            return `<td id="${id}">0</td>`;
        }).join("");
        table.appendChild(tbody);
        table.appendChild(countRow);

        groupStudents.forEach(name => {
            monthMap[monthKey].forEach(w => {
                const week = w.week;
                setTimeout(() => {
                    const checkbox = document.querySelector(`input[name='${name}-w${week}']`);
                    const countId = `count-${groupName.replace(/\s+/g, '')}-w${week}`;
                    checkbox.addEventListener("change", () => {
                        const countCell = document.getElementById(countId);
                        const groupBoxes = groupStudents.map(n => document.querySelector(`input[name='${n}-w${week}']`)).filter(Boolean);
                        countCell.textContent = groupBoxes.filter(b => b.checked).length;
                    });
                }, 0);
            });
        });

        group.appendChild(table);
        container.appendChild(group);
    }

    const totalRow = document.createElement("div");
    totalRow.className = "group";
    totalRow.innerHTML = "<h2>全部總出席</h2>";
    const totalTable = document.createElement("table");
    const totalHead = document.createElement("thead");
    const totalHeadRow = document.createElement("tr");
    totalHeadRow.innerHTML = "<th>週次</th>" + monthMap[monthKey].map(w => {
        const d = new Date(w.date);
        return `<th>W${w.week} - ${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日</th>`;
    }).join("");
    totalHead.appendChild(totalHeadRow);
    totalTable.appendChild(totalHead);

    const totalBody = document.createElement("tbody");
    const totalCountRow = document.createElement("tr");
    totalCountRow.innerHTML = "<td><strong>出席人數</strong></td>" + monthMap[monthKey].map(w => {
        const id = `total-week-${w.week}`;
        return `<td id="${id}">0</td>`;
    }).join("");
    totalBody.appendChild(totalCountRow);
    totalTable.appendChild(totalBody);
    totalRow.appendChild(totalTable);
    container.insertBefore(totalRow, container.firstChild);

    monthMap[monthKey].forEach(w => {
        const week = w.week;
        setTimeout(() => {
            const totalId = `total-week-${week}`;
            const updateTotal = () => {
                const allWeekBoxes = document.querySelectorAll(`input[name$='-w${week}']`);
                const checked = Array.from(allWeekBoxes).filter(b => b.checked).length;
                const totalCell = document.getElementById(totalId);
                if (totalCell) totalCell.textContent = checked;
            };
            document.querySelectorAll(`input[name$='-w${week}']`).forEach(box => {
                box.addEventListener("change", updateTotal);
            });
        }, 0);
    });

    tablesContainer.appendChild(container);
});

// =========================
// Save checkbox states to localStorage
// =========================
document.addEventListener("DOMContentLoaded", () => {
    const checkboxes = document.querySelectorAll("input[type='checkbox']");
    checkboxes.forEach(box => {
        const key = box.name;
        box.checked = localStorage.getItem(key) === "true";
        box.addEventListener("change", () => {
            localStorage.setItem(key, box.checked);
        });
    });

    // Trigger attendance calculations
    for (let i = 1; i <= weeks; i++) {
        const event = new Event("change");
        document.querySelectorAll(`input[name$='-w${i}']`).forEach(box => box.dispatchEvent(event));
    }

    // Draw Chart
    
    const chartCanvas = document.createElement("canvas");
    chartCanvas.id = "attendanceChart";
    // Full width chart
    chartCanvas.style.width = "100%";
    chartCanvas.style.height = "400px";
    chartCanvas.style.display = "block";
    const chartContainer = document.getElementById("chartContainer");
    chartContainer.appendChild(chartCanvas);
setTimeout(() => {
        const weeklyTotals = Array.from({ length: weeks }, (_, i) => {
            const week = i + 1;
            const allBoxes = document.querySelectorAll(`input[name$='-w${week}']`);
            return Array.from(allBoxes).filter(b => b.checked).length;
        });

        const ctx = document.getElementById("attendanceChart").getContext("2d");
        new Chart(ctx, {
            type: "line",
            data: {
                labels: weeklyTotals.map((_, i) => `W${i + 1}`),
                datasets: [{
                    label: "每週總出席人數",
                    data: weeklyTotals,
                    borderWidth: 2,
                    fill: false,
                    tension: 0.2
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '出席人數'
                        }
                    }
                }
            }
        });
    }, 500);
});
