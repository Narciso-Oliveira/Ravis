"use client";

import { useEffect, useState } from "react";

type Cell = string | number | boolean | null;
type Row = Cell[];

export default function Home() {
  const [rows, setRows] = useState<Row[]>([]);
  const [sheetName, setSheetName] = useState("Planilha");
  const [sha, setSha] = useState("");
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  async function loadSheet() {
    setLoading(true);

    const response = await fetch("/api/sheet");
    const data = await response.json();

    if (!response.ok) {
      alert(data.error || "Erro ao carregar planilha.");
      setLoading(false);
      return;
    }

    setRows(data.rows);
    setSheetName(data.sheetName);
    setSha(data.sha);
    setLoading(false);
  }

  function updateCell(rowIndex: number, colIndex: number, value: string) {
    setRows((currentRows) => {
      const nextRows = currentRows.map((row) => [...row]);

      while (!nextRows[rowIndex]) {
        nextRows.push([]);
      }

      nextRows[rowIndex][colIndex] = value;

      return nextRows;
    });
  }

  function addRow() {
    const columnCount = Math.max(...rows.map((row) => row.length), 1);
    setRows([...rows, Array(columnCount).fill("")]);
  }

  function addColumn() {
    setRows(rows.map((row) => [...row, ""]));
  }

  async function saveSheet() {
    setSaving(true);

    const response = await fetch("/api/sheet", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ rows, sheetName, sha }),
    });

    const data = await response.json();

    setSaving(false);

    if (!response.ok) {
      alert(data.error || "Erro ao salvar planilha.");
      return;
    }

    alert("Planilha salva no GitHub com sucesso.");
    await loadSheet();
  }

  useEffect(() => {
    loadSheet();
  }, []);

  if (loading) {
    return <main style={{ padding: 24 }}>Carregando planilha...</main>;
  }

  const columnCount = Math.max(...rows.map((row) => row.length), 1);

  return (
    <main style={{ padding: 24, fontFamily: "Arial, sans-serif" }}>
      <h1>Editor da planilha Ravis</h1>

      <div style={{ marginBottom: 16, display: "flex", gap: 8 }}>
        <button onClick={addRow}>Adicionar linha</button>
        <button onClick={addColumn}>Adicionar coluna</button>
        <button onClick={saveSheet} disabled={saving}>
          {saving ? "Salvando..." : "Salvar no GitHub"}
        </button>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse" }}>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {Array.from({ length: columnCount }).map((_, colIndex) => (
                  <td
                    key={colIndex}
                    style={{
                      border: "1px solid #ccc",
                      padding: 4,
                    }}
                  >
                    <input
                      value={String(row[colIndex] ?? "")}
                      onChange={(event) =>
                        updateCell(rowIndex, colIndex, event.target.value)
                      }
                      style={{
                        width: 160,
                        border: "none",
                        outline: "none",
                      }}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </main>
  );
}