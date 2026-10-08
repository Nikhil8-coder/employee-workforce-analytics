"""
Standalone HTML Web App Generator
Exports the entire Employee Analytics suite as a zero-dependency interactive HTML5/Tailwind/Chart.js
web application that runs in any browser and can be hosted directly on GitHub Pages, Netlify, or Vercel.
"""

import os
import json
import pandas as pd

def build_standalone_dashboard(csv_path: str, output_html: str):
    df = pd.read_csv(csv_path)
    records_json = df.to_json(orient="records")

    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Enterprise Employee Intelligence & Analytics Web Platform</title>
  <!-- Tailwind CSS & Chart.js CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
</head>
<body class="bg-slate-50 text-slate-800 antialiased font-sans min-h-screen">

  <!-- Top Navbar -->
  <header class="bg-indigo-900 text-white shadow-md sticky top-0 z-50">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
      <div class="flex items-center space-x-3">
        <i class="fa-solid fa-chart-line text-indigo-400 text-2xl"></i>
        <div>
          <h1 class="font-bold text-lg leading-tight">Workforce Intelligence & Compensation Suite</h1>
          <p class="text-xs text-indigo-300">Production Analytics &bull; Live Client-Side Engine</p>
        </div>
      </div>
      <div class="flex items-center space-x-3">
        <span class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
          <span class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse mr-1.5"></span> Live Engine
        </span>
        <button onclick="exportCSV()" class="bg-indigo-700 hover:bg-indigo-600 px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition">
          <i class="fa-solid fa-download"></i> Export Filtered CSV
        </button>
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">

    <!-- Filters Bar -->
    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
      <div class="flex items-center justify-between mb-4">
        <h2 class="font-bold text-slate-800 flex items-center gap-2">
          <i class="fa-solid fa-sliders text-indigo-600"></i> Interactive Controls & Multi-Filters
        </h2>
        <button onclick="resetFilters()" class="text-xs font-semibold text-indigo-600 hover:text-indigo-800">
          Reset All Filters
        </button>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
        <!-- Search -->
        <div>
          <label class="block text-xs font-semibold text-slate-500 mb-1">Search Name, Role, or ID</label>
          <div class="relative">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
            <input type="text" id="searchInput" oninput="applyFilters()" placeholder="e.g. Priya or Engineer..." class="w-full pl-8 pr-3 py-2 text-sm border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
          </div>
        </div>

        <!-- Department -->
        <div>
          <label class="block text-xs font-semibold text-slate-500 mb-1">Department</label>
          <select id="deptSelect" onchange="applyFilters()" class="w-full px-3 py-2 text-sm border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none">
            <option value="ALL">All Departments</option>
          </select>
        </div>

        <!-- Salary Threshold -->
        <div>
          <div class="flex justify-between items-center mb-1">
            <label class="text-xs font-semibold text-slate-500">Min Salary (Threshold)</label>
            <span id="salaryVal" class="text-xs font-bold text-indigo-600">₹0</span>
          </div>
          <input type="range" id="salaryRange" min="0" max="350000" step="5000" value="0" oninput="updateSalaryLabel(); applyFilters();" class="w-full accent-indigo-600 cursor-pointer">
        </div>

        <!-- Experience Range -->
        <div>
          <div class="flex justify-between items-center mb-1">
            <label class="text-xs font-semibold text-slate-500">Min Experience</label>
            <span id="expVal" class="text-xs font-bold text-indigo-600">0 Years</span>
          </div>
          <input type="range" id="expRange" min="0" max="25" step="1" value="0" oninput="updateExpLabel(); applyFilters();" class="w-full accent-indigo-600 cursor-pointer">
        </div>
      </div>
    </div>

    <!-- KPI Metric Cards -->
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
        <p class="text-xs uppercase font-semibold text-slate-400 tracking-wider">Filtered Headcount</p>
        <p id="kpiHeadcount" class="text-3xl font-extrabold text-slate-900 mt-1">0</p>
        <p id="kpiHeadcountPct" class="text-xs text-indigo-600 mt-1 font-medium">100% of workforce</p>
      </div>

      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
        <p class="text-xs uppercase font-semibold text-slate-400 tracking-wider">Total Annual Payroll</p>
        <p id="kpiPayroll" class="text-3xl font-extrabold text-indigo-600 mt-1">₹0</p>
        <p class="text-xs text-slate-400 mt-1">Cumulative Base</p>
      </div>

      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
        <p class="text-xs uppercase font-semibold text-slate-400 tracking-wider">Average Salary</p>
        <p id="kpiAvgSalary" class="text-3xl font-extrabold text-slate-900 mt-1">₹0</p>
        <p id="kpiMedianSalary" class="text-xs text-slate-500 mt-1">Median: ₹0</p>
      </div>

      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
        <p class="text-xs uppercase font-semibold text-slate-400 tracking-wider">Salary Range</p>
        <p id="kpiSalaryRange" class="text-2xl font-extrabold text-emerald-600 mt-1">₹0 - ₹0</p>
        <p id="kpiAvgExp" class="text-xs text-slate-500 mt-1">Avg Exp: 0 yrs</p>
      </div>
    </div>

    <!-- Charts Row -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
        <h3 class="font-bold text-slate-800 text-sm mb-4 flex items-center gap-2">
          <i class="fa-solid fa-chart-column text-indigo-600"></i> Departmental Headcount & Average Compensation
        </h3>
        <div class="h-64">
          <canvas id="deptChart"></canvas>
        </div>
      </div>

      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-5">
        <h3 class="font-bold text-slate-800 text-sm mb-4 flex items-center gap-2">
          <i class="fa-solid fa-chart-pie text-indigo-600"></i> Compensation Distribution Brackets
        </h3>
        <div class="h-64">
          <canvas id="salaryDistChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Data Table -->
    <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
      <div class="p-5 border-b border-slate-100 flex items-center justify-between">
        <div>
          <h3 class="font-bold text-slate-900">Workforce Cohort Roster</h3>
          <p class="text-xs text-slate-500 mt-0.5" id="tableRecordCount">Showing 0 records</p>
        </div>
        <div class="text-xs font-semibold text-slate-400">
          Click headers to sort
        </div>
      </div>
      <div class="overflow-x-auto max-h-[500px]">
        <table class="min-w-full divide-y divide-slate-100 text-sm text-left">
          <thead class="bg-slate-50 sticky top-0 text-slate-600 font-semibold text-xs tracking-wider uppercase">
            <tr>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Employee_ID')">ID</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Name')">Name</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Department')">Department</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Job_Title')">Role</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Salary')">Salary (₹)</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Experience_Years')">Experience</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Performance_Score')">Rating</th>
              <th class="py-3 px-4 cursor-pointer hover:bg-slate-100" onclick="sortTable('Location')">Location</th>
            </tr>
          </thead>
          <tbody id="tableBody" class="divide-y divide-slate-100 text-slate-700">
            <!-- Rendered dynamically -->
          </tbody>
        </table>
      </div>
    </div>

  </main>

  <!-- Embed data into frontend script -->
  <script>
    const RAW_DATA = {records_json};
    let currentData = [...RAW_DATA];
    let sortCol = "Salary";
    let sortAsc = false;
    let deptChart = null;
    let salaryDistChart = null;

    // Initialize Dropdowns & Dashboard
    function init() {{
      const depts = [...new Set(RAW_DATA.map(r => r.Department))].sort();
      const sel = document.getElementById("deptSelect");
      depts.forEach(d => {{
        const opt = document.createElement("option");
        opt.value = d;
        opt.textContent = d;
        sel.appendChild(opt);
      }});

      applyFilters();
    }}

    function updateSalaryLabel() {{
      const val = Number(document.getElementById("salaryRange").value);
      document.getElementById("salaryVal").textContent = "₹" + val.toLocaleString();
    }}

    function updateExpLabel() {{
      const val = Number(document.getElementById("expRange").value);
      document.getElementById("expVal").textContent = val + " Years";
    }}

    function resetFilters() {{
      document.getElementById("searchInput").value = "";
      document.getElementById("deptSelect").value = "ALL";
      document.getElementById("salaryRange").value = 0;
      document.getElementById("expRange").value = 0;
      updateSalaryLabel();
      updateExpLabel();
      applyFilters();
    }}

    function applyFilters() {{
      const search = document.getElementById("searchInput").value.toLowerCase().trim();
      const dept = document.getElementById("deptSelect").value;
      const minSalary = Number(document.getElementById("salaryRange").value);
      const minExp = Number(document.getElementById("expRange").value);

      currentData = RAW_DATA.filter(r => {{
        if (dept !== "ALL" && r.Department !== dept) return false;
        if (r.Salary < minSalary) return false;
        if (r.Experience_Years < minExp) return false;
        if (search) {{
          const matchName = String(r.Name || "").toLowerCase().includes(search);
          const matchRole = String(r.Job_Title || "").toLowerCase().includes(search);
          const matchId = String(r.Employee_ID || "").toLowerCase().includes(search);
          if (!matchName && !matchRole && !matchId) return false;
        }}
        return true;
      }});

      sortData();
      renderKPIs();
      renderCharts();
      renderTable();
    }}

    function sortTable(col) {{
      if (sortCol === col) {{
        sortAsc = !sortAsc;
      }} else {{
        sortCol = col;
        sortAsc = true;
      }}
      sortData();
      renderTable();
    }}

    function sortData() {{
      currentData.sort((a, b) => {{
        let v1 = a[sortCol];
        let v2 = b[sortCol];
        if (typeof v1 === "string") {{
          return sortAsc ? v1.localeCompare(v2) : v2.localeCompare(v1);
        }}
        return sortAsc ? v1 - v2 : v2 - v1;
      }});
    }}

    function renderKPIs() {{
      const count = currentData.length;
      document.getElementById("kpiHeadcount").textContent = count.toLocaleString();
      document.getElementById("kpiHeadcountPct").textContent = ((count / RAW_DATA.length) * 100).toFixed(1) + "% of total workforce";

      if (count === 0) {{
        document.getElementById("kpiPayroll").textContent = "₹0";
        document.getElementById("kpiAvgSalary").textContent = "₹0";
        document.getElementById("kpiMedianSalary").textContent = "Median: ₹0";
        document.getElementById("kpiSalaryRange").textContent = "₹0 - ₹0";
        document.getElementById("kpiAvgExp").textContent = "Avg Exp: 0 yrs";
        return;
      }}

      const totalPayroll = currentData.reduce((acc, r) => acc + r.Salary, 0);
      const avgSalary = Math.round(totalPayroll / count);

      const sortedSalaries = [...currentData.map(r => r.Salary)].sort((a, b) => a - b);
      const medianSalary = sortedSalaries[Math.floor(count / 2)];
      const minSalary = sortedSalaries[0];
      const maxSalary = sortedSalaries[count - 1];

      const avgExp = (currentData.reduce((acc, r) => acc + (r.Experience_Years || 0), 0) / count).toFixed(1);

      document.getElementById("kpiPayroll").textContent = "₹" + totalPayroll.toLocaleString();
      document.getElementById("kpiAvgSalary").textContent = "₹" + avgSalary.toLocaleString();
      document.getElementById("kpiMedianSalary").textContent = "Median: ₹" + medianSalary.toLocaleString();
      document.getElementById("kpiSalaryRange").textContent = "₹" + minSalary.toLocaleString() + " - ₹" + maxSalary.toLocaleString();
      document.getElementById("kpiAvgExp").textContent = "Avg Tenure: " + avgExp + " yrs";
    }}

    function renderCharts() {{
      // Department breakdown
      const deptMap = {{}};
      currentData.forEach(r => {{
        if (!deptMap[r.Department]) deptMap[r.Department] = {{ count: 0, totalSal: 0 }};
        deptMap[r.Department].count += 1;
        deptMap[r.Department].totalSal += r.Salary;
      }});

      const deptLabels = Object.keys(deptMap);
      const deptCounts = deptLabels.map(d => deptMap[d].count);
      const deptAvgs = deptLabels.map(d => Math.round(deptMap[d].totalSal / deptMap[d].count));

      if (deptChart) deptChart.destroy();
      const ctx1 = document.getElementById("deptChart").getContext("2d");
      deptChart = new Chart(ctx1, {{
        type: 'bar',
        data: {{
          labels: deptLabels,
          datasets: [{{
            label: 'Headcount',
            data: deptCounts,
            backgroundColor: '#4F46E5',
            borderRadius: 6
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{ legend: {{ display: false }} }}
        }}
      }});

      // Salary distribution brackets
      const brackets = {{
        "< ₹75K": 0,
        "₹75K - ₹125K": 0,
        "₹125K - ₹200K": 0,
        "₹200K - ₹300K": 0,
        "> ₹300K": 0
      }};

      currentData.forEach(r => {{
        if (r.Salary < 75000) brackets["< ₹75K"]++;
        else if (r.Salary <= 125000) brackets["₹75K - ₹125K"]++;
        else if (r.Salary <= 200000) brackets["₹125K - ₹200K"]++;
        else if (r.Salary <= 300000) brackets["₹200K - ₹300K"]++;
        else brackets["> ₹300K"]++;
      }});

      if (salaryDistChart) salaryDistChart.destroy();
      const ctx2 = document.getElementById("salaryDistChart").getContext("2d");
      salaryDistChart = new Chart(ctx2, {{
        type: 'doughnut',
        data: {{
          labels: Object.keys(brackets),
          datasets: [{{
            data: Object.values(brackets),
            backgroundColor: ['#6366F1', '#3B82F6', '#10B981', '#F59E0B', '#EF4444']
          }}]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false
        }}
      }});
    }}

    function renderTable() {{
      const tbody = document.getElementById("tableBody");
      tbody.innerHTML = "";
      document.getElementById("tableRecordCount").textContent = "Showing " + currentData.length + " employee records";

      const preview = currentData.slice(0, 100);
      preview.forEach(r => {{
        const tr = document.createElement("tr");
        tr.className = "hover:bg-slate-50 transition";
        tr.innerHTML = `
          <td class="py-3 px-4 font-mono text-xs font-semibold text-slate-500">${{r.Employee_ID}}</td>
          <td class="py-3 px-4 font-semibold text-slate-900">${{r.Name}}</td>
          <td class="py-3 px-4"><span class="px-2 py-0.5 rounded-full text-xs font-medium bg-indigo-50 text-indigo-700">${{r.Department}}</span></td>
          <td class="py-3 px-4 text-xs text-slate-600">${{r.Job_Title || '-'}}</td>
          <td class="py-3 px-4 font-semibold text-slate-900">₹${{r.Salary.toLocaleString()}}</td>
          <td class="py-3 px-4 text-xs text-slate-600">${{r.Experience_Years || 0}} yrs</td>
          <td class="py-3 px-4"><span class="text-xs font-semibold text-amber-600">★ ${{r.Performance_Score || 3.0}}</span></td>
          <td class="py-3 px-4 text-xs text-slate-500">${{r.Location || '-'}}</td>
        `;
        tbody.appendChild(tr);
      }});
    }}

    function exportCSV() {{
      if (currentData.length === 0) return alert("No records to export.");
      const headers = Object.keys(currentData[0]);
      const csvRows = [headers.join(",")];

      currentData.forEach(r => {{
        const values = headers.map(h => {{
          const val = r[h] === undefined ? "" : String(r[h]).replace(/"/g, '""');
          return `"${{val}}"`;
        }});
        csvRows.push(values.join(","));
      }});

      const blob = new Blob([csvRows.join("\\n")], {{ type: "text/csv;charset=utf-8;" }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "workforce_filtered_export.csv";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }}

    window.onload = init;
  </script>
</body>
</html>
"""
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_template)
    print(f"Standalone Web App successfully generated at: {output_html}")

if __name__ == "__main__":
    csv_file = os.path.join(os.path.dirname(__file__), "data", "employees.csv")
    out_file = os.path.join(os.path.dirname(__file__), "index.html")
    build_standalone_dashboard(csv_file, out_file)
