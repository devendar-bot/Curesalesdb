const URL = import.meta.env.VITE_API_URL;
const TOKEN = import.meta.env.VITE_API_TOKEN;

async function parse(r) {
  let d;
  try { d = await r.json(); } catch { throw new Error("Backend did not return JSON – check VITE_API_URL and that the Web app is deployed for 'Anyone'"); }
  if (d.error) throw new Error(d.error);
  return d;
}

// text/plain avoids a CORS pre-flight, which Apps Script cannot answer.
const post = (payload) =>
  fetch(URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ token: TOKEN || "", ...payload }) }).then(parse);

// Read every tab straight from the Google Sheet.
export async function loadAll() {
  if (!URL || URL.includes("PASTE_")) throw new Error("VITE_API_URL is not set – paste your Apps Script Web app URL into .env / Vercel");
  return parse(await fetch(`${URL}?action=all&token=${encodeURIComponent(TOKEN || "")}`));
}

// Write mapped rows into one source tab of the Sheet.
export const importRows = (tab, rows, mode) => post({ action: "import", tab, mode, rows });

// Write "Qty to be ordered" + "Status" into the Transaction Data tab.
export const saveOrder = (items, vars) => post({ action: "saveOrder", items, vars });
