
const sheetURL = "https://script.google.com/macros/s/AKfycbyLKxAMq7_w5VKbLfbqspwZ2JdKuqagql7xY3y_JpdEF1R8CvQWV0IepFCPCjrx9wJx/exec";
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
    if (!monthMap[label]) monthMap[label] = [];
    monthMap[label].push({ week: i + 1, date: date.toISOString().split("T")[0] });
}

const monthTabs = document.getElementById("monthTabs");
const tablesContainer = document.getElementById("tablesContainer");

Object.keys(monthMap).forEach((monthKey, idx) => {
    const tab = document.createElement("div");
    tab.className = "tab" + (idx === 0 ? " active" : "");
    tab.innerText = monthKey;
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

        headRow.innerHTML = `<th>姓名</th>` + monthMap[monthKey].map(w => {
            const d = new Date(w.date);
            return `<th>W${w.week} - ${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日</th>`;
        }).join("");
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        groupStudents.forEach(name => {
            const row = document.createElement("tr");
            row.innerHTML = `<td>${name}</td>` + monthMap[monthKey].map(w =>
                `<td><input type="checkbox" name="${name}-w${w.week}"></td>`).join("");
            tbody.appendChild(row);
        });
        table.appendChild(tbody);
        group.appendChild(table);
        container.appendChild(group);
    }

    tablesContainer.appendChild(container);
});

// Sync button
document.getElementById("sync").onclick = () => {
    fetch(sheetURL)
        .then(res => res.json())
        .then(data => {
            data.forEach(entry => {
                const name = entry.Student;
                const week = entry.Week.replace("W", "");
                const checkbox = document.querySelector(`input[name="${name}-w${week}"]`);
                if (checkbox) {
                    checkbox.checked = entry.Attended == "1";
                    localStorage.setItem(`${name}-w${week}`, checkbox.checked);
                }
            });
            alert("同步完成！");
        }).catch(err => {
            console.error(err);
            alert("同步失敗");
        });
};

// Export button
document.getElementById("export").onclick = () => {
    const rows = [["Group", "Student", "Week", "Date", "Attended"]];
    Object.entries(students).forEach(([group, members]) => {
        members.forEach(student => {
            for (let i = 1; i <= weeks; i++) {
                const cb = document.querySelector(`input[name="${student}-w${i}"]`);
                if (cb) {
                    const date = new Date(startDate);
                    date.setDate(date.getDate() + (i - 1) * 7);
                    rows.push([group, student, `W${i}`, date.toISOString().split("T")[0], cb.checked ? "1" : "0"]);
                }
            }
        });
    });
    const csv = rows.map(r => r.join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "attendance.csv";
    link.click();
};
