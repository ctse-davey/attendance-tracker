
const sheetURL = "https://script.google.com/macros/s/AKfycbyLKxAMq7_w5VKbLfbqspwZ2JdKuqagql7xY3y_JpdEF1R8CvQWV0IepFCPCjrx9wJx/exec";
const startDate = new Date("2025-04-21");
const weeks = 49;
const students = {
    "第1組 - 碧蓮 ★ 文信": ["碧蓮", "文信", "焦華彬", "黃潤萍", "胡灝文", "羅青玲", "廖喜兒", "任麥寶燕", "羅美好", "盧寶萍BoBo"],
    "第2組 - 桂蘭 ★ Jenny": ["桂蘭", "Jenny", "蕭慧明", "周貴珍", "林寶珍", "吳建容", "梁漪好", "黎桂英", "方娟梅"],
    "第3組 - 彭姑娘 ★ Winnie": ["彭姑娘", "Winnie", "何天嬌", "溫房嬌", "朱寶兒", "李玉華", "符志蓮Lilian", "蔡麗萍Doris", "鄭瑞琼Gloria", "黃綺華"]
};

// The rest of the logic (cleaned)
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

        headRow.innerHTML = `<th>姓名</th>` + monthMap[monthKey].map(w => {
        const d = new Date(w.date);
        const formattedDate = `W${w.week} - ${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`;
        return `<th>${formattedDate}</th>`;
    }).join("");
        thead.appendChild(headRow);
        table.appendChild(thead);

        const tbody = document.createElement("tbody");
        groupStudents.forEach(name => {
            const row = document.createElement("tr");
            row.innerHTML = `<td>${name}</td>` + monthMap[monthKey].map(w => `<td><input type="checkbox" name="${name}-w${w.week}"></td>`).join("");
            tbody.appendChild(row);
        });

        
        const countRow = document.createElement("tr");
        countRow.innerHTML = `<td><strong>出席人數</strong></td>` + monthMap[monthKey].map(w => {
            const id = `count-${groupName.replace(/\s+/g, '')}-w${w.week}`;
            return `<td id="${id}">0</td>`;
        }).join("");
        table.appendChild(tbody);
        table.appendChild(countRow);

        // Add checkbox listeners
        groupStudents.forEach(name => {
            monthMap[monthKey].forEach(w => {
                const week = w.week;
                setTimeout(() => {
                    const checkbox = document.querySelector(`input[name='${name}-w${week}']`);
                    const countId = `count-${groupName.replace(/\s+/g, '')}-w${week}`;
                    checkbox.addEventListener('change', () => {
                        const countCell = document.getElementById(countId);
                        const allBoxes = document.querySelectorAll(`input[name$='-w${week}']`);
                        const checkedCount = Array.from(allBoxes).filter(b => b.checked).length;
                        countCell.textContent = checkedCount;
                    });
                }, 0);
            });
        });
    
        group.appendChild(table);
        container.appendChild(group);
    }

    
    // Add total attendance row across all groups
    const totalRow = document.createElement("div");
    totalRow.className = "group";
    totalRow.innerHTML = `<h2>全部總出席</h2>`;

    const totalTable = document.createElement("table");
    const totalHead = document.createElement("thead");
    const totalHeadRow = document.createElement("tr");

    totalHeadRow.innerHTML = `<th>週次</th>` + monthMap[monthKey].map(w => {
        const d = new Date(w.date);
        return `<th>W${w.week} - ${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日</th>`;
    }).join("");

    totalHead.appendChild(totalHeadRow);
    totalTable.appendChild(totalHead);

    const totalBody = document.createElement("tbody");
    const totalCountRow = document.createElement("tr");
    totalCountRow.innerHTML = `<td><strong>出席人數</strong></td>` + monthMap[monthKey].map(w => {
        const id = `total-week-${w.week}`;
        return `<td id="${id}">0</td>`;
    }).join("");

    totalBody.appendChild(totalCountRow);
    totalTable.appendChild(totalBody);
    totalRow.appendChild(totalTable);
    container.appendChild(totalRow);

    // Live update total counts across all groups
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

