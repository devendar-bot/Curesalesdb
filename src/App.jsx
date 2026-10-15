import React, { useEffect, useState } from "react";
import Dashboard from "./Dashboard.jsx";
import { loadAll, importRows, saveOrder } from "./api.js";

const SHEET_ID = import.meta.env.VITE_SHEET_ID;
const SHEET_URL = SHEET_ID ? `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit` : null;

export default function App() {
  const [data, setData] = useState(null);
  const [err, setErr] = useState(null);

  useEffect(() => { loadAll().then(setData).catch((e) => setErr(e.message)); }, []);

  // Each write goes to the Google Sheet first; then everything is re-read so the numbers recalculate.
  const onImport = async (tab, rows, mode, setT) => { await importRows(tab, rows, mode); setT(await loadAll()); };
  const onSave = async (items, vars, setT) => { await saveOrder(items, vars); setT(await loadAll()); };
  const onReload = async (setT) => setT(await loadAll());

  if (err) return <p style={{ color: "crimson", padding: 16 }}>Could not load the Google Sheet: {err}</p>;
  if (!data) return <p style={{ padding: 16 }}>Loading data from Google Sheet…</p>;
  return <Dashboard initial={data} onImport={onImport} onSave={onSave} onReload={onReload} sheetUrl={SHEET_URL} />;
}
