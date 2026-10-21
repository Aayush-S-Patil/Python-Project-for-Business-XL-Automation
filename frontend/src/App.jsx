
import { useState } from "react";
import * as XLSX from "xlsx";

import {
  BarChart, Bar,
  PieChart, Pie, Cell,
  LineChart, Line,
  XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer
} from "recharts";

import {
  Upload, FileSpreadsheet, Download, RefreshCw,
  CheckCircle, AlertCircle, Percent
} from "lucide-react";
import "./App.css";

const API_URL = "http://127.0.0.1:5000/api/process";

export default function App() {
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");

  async function handleFile(selectedFile) {
    setError("");
    setSuccess("");
    setRows([]);
    setFile(null);

    if (!selectedFile) return;

    if (!selectedFile.name.toLowerCase().endsWith(".xlsx")) {
      setError("Please select an .xlsx Excel file.");
      return;
    }

    try {
      const buffer = await selectedFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets["Sheet1"];

      if (!sheet) {
        setError('The workbook must contain a sheet named "Sheet1".');
        return;
      }

      const data = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: null
      });


      const headers = (data[0] || []).map(value =>
        String(value ?? "").trim().toLowerCase()
      );

      const categoryIndex = headers.findIndex(header =>
        /^(expense )?category$|^category name$|^expense type$|^type$/.test(header)
      );


      const preview = [];

      for (let i = 1; i < data.length; i++) {
        const original = data[i]?.[2];

        if (original === null || original === undefined || original === "") {
          continue;
        }

        if (typeof original !== "number" || !Number.isFinite(original)) {
          setError(`Invalid numeric price in cell C${i + 1}.`);
          return;
        }

        preview.push({
          row: i + 1,
          name: data[i]?.[0] ?? `Row ${i + 1}`,
          original,
          discounted: Math.round((original * 0.9 + Number.EPSILON) * 100) / 100,
          category:
            categoryIndex >= 0 && data[i]?.[categoryIndex] != null
              ? String(data[i][categoryIndex])
              : String(data[i]?.[0] ?? `Row ${i + 1}`),
        });
      }

      if (!preview.length) {
        setError("No prices found in column C, starting at row 2.");
        return;
      }

      setFile(selectedFile);
      setRows(preview);
    } catch {
      setError("Unable to read this Excel file.");
    }
  }

  async function processFile() {
    if (!file) return;

    setLoading(true);
    setError("");
    setSuccess("");

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch(API_URL, {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        const result = await response.json();
        throw new Error(result.error || "Excel processing failed.");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");

      link.href = url;
      link.download = "processed_prices.xlsx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      setSuccess(
        `Successfully processed ${rows.length} price rows. Your Excel file has been downloaded.`
      );
    } catch (err) {
      setError(
        err.message || "Could not connect to the Python backend. Is it running?"
      );
    } finally {
      setLoading(false);
    }
  }

  function reset() {
    setFile(null);
    setRows([]);
    setError("");
    setSuccess("");
  }


  const CHART_COLORS = [
    "#FF6B6B",
    "#4ECDC4",
    "#A78BFA",
    "#FFC857",
    "#60A5FA",
    "#F472B6",
    "#86C96F",
    "#FB923C",
  ];

  const categoryTotals = Object.values(
    rows.reduce((groups, row) => {
      const category = row.category || row.name;
      if (!groups[category]) {
        groups[category] = { name: category, value: 0 };
      }
      groups[category].value += row.original;
      return groups;
    }, {})
  ).sort((a, b) => b.value - a.value);

  const savingsData = rows.map(row => ({
    name: String(row.name).slice(0, 18),
    row: row.row,
    savings: Number((row.original - row.discounted).toFixed(2)),
  }));

  const trendData = rows.map(row => ({
    row: row.row,
    original: row.original,
    discounted: row.discounted,
  }));


  return (
    <main className="app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-icon"><FileSpreadsheet /></div>
          <div>
            <h1>ExcelPrice</h1>
            <p>Smart spreadsheet processing</p>
          </div>
        </div>
        <span className="status"><span /> Ready to process</span>
      </header>

      <section className="hero">
        <span className="eyebrow"><Percent size={15} /> EXCEL AUTOMATION</span>
        <h2>Update prices.<br /><span>Save time.</span></h2>
        <p>
          Upload your spreadsheet, apply a 10% discount to every price,
          compare results, and download an Excel workbook with a chart.
        </p>
      </section>

      <section className="upload-card">
        <div className="section-heading">
          <div>
            <h3>Upload your workbook</h3>
            <p>Choose an Excel .xlsx file with prices in column C.</p>
          </div>
          <FileSpreadsheet size={28} className="muted-icon" />
        </div>

        <label className="dropzone">
          <div className="upload-icon"><Upload size={24} /></div>
          <strong>{file ? file.name : "Click to select your Excel file"}</strong>
          <span>{file ? `${(file.size / 1024).toFixed(1)} KB` : "XLSX format · Maximum 25 MB"}</span>
          <input
            type="file"
            accept=".xlsx"
            onChange={e => handleFile(e.target.files?.[0])}
          />
        </label>

        {file && (
          <div className="file-info">
            <CheckCircle size={18} />
            <span>{rows.length} populated price rows detected</span>
          </div>
        )}

        <div className="actions">
          <button
            className="primary"
            onClick={processFile}
            disabled={!file || loading}
          >
            {loading ? <RefreshCw className="spin" size={18} /> : <Download size={18} />}
            {loading ? "Processing..." : "Process & Download"}
          </button>
          <button className="secondary" onClick={reset} disabled={loading}>
            Reset
          </button>
        </div>

        {error && <div className="message error"><AlertCircle size={18} />{error}</div>}
        {success && <div className="message success"><CheckCircle size={18} />{success}</div>}
      </section>

      {rows.length > 0 && (
        <>
          <section className="stats">
            <div className="stat-card">
              <span>Rows processed</span>
              <strong>{rows.length.toLocaleString()}</strong>
              <small>Automatically detected</small>
            </div>
            <div className="stat-card">
              <span>Discount applied</span>
              <strong>10%</strong>
              <small>On each valid price</small>
            </div>
            <div className="stat-card">
              <span>Total savings</span>
              <strong>
                {rows.reduce((sum, r) => sum + r.original - r.discounted, 0)
                  .toLocaleString(undefined, { maximumFractionDigits: 2 })}
              </strong>
              <small>Across previewed rows</small>
            </div>
          </section>


          <section className="panel">
            <div className="section-heading">
              <div>
                <h3>Price comparison</h3>
                <p>Original prices versus prices after discount.</p>
              </div>
            </div>

            <div className="chart" style={{ background: "#F8FAFC", borderRadius: "12px", padding: "12px" }}>
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={rows}
                  margin={{ top: 15, right: 20, left: 10, bottom: 20 }}
                >
                  <CartesianGrid
                    strokeDasharray="4 4"
                    stroke="#DCE3ED"
                    vertical={false}
                  />

                  <XAxis
                    dataKey="row"
                    stroke="#475569"
                    tick={{ fill: "#475569", fontSize: 11 }}
                    tickLine={{ stroke: "#CBD5E1" }}
                    axisLine={{ stroke: "#CBD5E1" }}
                    label={{
                      value: "Excel row",
                      position: "insideBottom",
                      offset: -10,
                      fill: "#334155",
                      fontSize: 12
                    }}
                  />

                  <YAxis
                    stroke="#475569"
                    tick={{ fill: "#475569", fontSize: 11 }}
                    tickLine={false}
                    axisLine={false}
                    width={65}
                  />

                  <Tooltip
                    formatter={(value, name) => [
                      Number(value).toLocaleString(undefined, {
                        maximumFractionDigits: 2
                      }),
                      name
                    ]}
                    labelFormatter={(label) => `Excel row ${label}`}
                    contentStyle={{
                      backgroundColor: "#FFFFFF",
                      border: "1px solid #CBD5E1",
                      borderRadius: "10px",
                      color: "#0F172A",
                      boxShadow: "0 4px 14px rgba(15, 23, 42, 0.10)"
                    }}
                    labelStyle={{
                      color: "#334155",
                      fontWeight: 600
                    }}
                  />

                  <Legend
                    wrapperStyle={{
                      paddingTop: "12px",
                      color: "#334155"
                    }}
                  />

                  <Bar
                    dataKey="original"
                    name="Original price"
                    fill="#F97316"
                    radius={[5, 5, 0, 0]}
                    maxBarSize={32}
                  />

                  <Bar
                    dataKey="discounted"
                    name="Discounted price"
                    fill="#14B8A6"
                    radius={[5, 5, 0, 0]}
                    maxBarSize={32}
                  />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>


          <section className="panel">
            <div className="section-heading">
              <div>
                <h3>Processed price preview</h3>
                <p>Showing up to 100 rows. All detected rows are processed.</p>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Excel row</th>
                    <th>Product / label</th>
                    <th>Original price</th>
                    <th>Discounted price</th>
                    <th>Saved</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.slice(0, 100).map(row => (
                    <tr key={row.row}>
                      <td>{row.row}</td>
                      <td>{String(row.name)}</td>
                      <td>{row.original.toLocaleString()}</td>
                      <td className="price">{row.discounted.toLocaleString()}</td>
                      <td className="saving">
                        {(row.original - row.discounted).toLocaleString(
                          undefined, { maximumFractionDigits: 2 }
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {rows.length > 100 && (
                <p className="table-note">
                  Preview limited to 100 rows. All {rows.length} rows are included in processing.
                </p>
              )}
            </div>
          </section>
        </>
      )}

      <footer>ExcelPrice · Automated Excel processing with Python + React</footer>
    </main>
  );
}
