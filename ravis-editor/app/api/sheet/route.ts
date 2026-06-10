import * as XLSX from "xlsx";

const owner = process.env.GITHUB_OWNER!;
const repo = process.env.GITHUB_REPO!;
const branch = process.env.GITHUB_BRANCH || "main";
const filePath = process.env.XLSX_PATH!;
const token = process.env.GITHUB_TOKEN!;

function githubHeaders() {
  return {
    Accept: "application/vnd.github+json",
    Authorization: `Bearer ${token}`,
    "X-GitHub-Api-Version": "2022-11-28",
  };
}

export async function GET() {
  const url =
    `https://api.github.com/repos/${owner}/${repo}/contents/${filePath}` +
    `?ref=${branch}`;

  const response = await fetch(url, {
    headers: githubHeaders(),
    cache: "no-store",
  });

  if (!response.ok) {
    return Response.json(
      { error: "Não foi possível ler a planilha no GitHub." },
      { status: response.status }
    );
  }

  const file = await response.json();

  const buffer = Buffer.from(file.content, "base64");

  const workbook = XLSX.read(buffer, { type: "buffer" });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  return Response.json({
    sheetName: firstSheetName,
    rows,
    sha: file.sha,
  });
}
export async function POST(request: Request) {
  const body = await request.json();

  const { rows, sheetName, sha } = body;

  if (!rows || !Array.isArray(rows)) {
    return Response.json(
      { error: "Dados inválidos." },
      { status: 400 }
    );
  }

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    sheetName || "Planilha"
  );

  const xlsxBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "buffer",
  });

  const base64Content = Buffer.from(xlsxBuffer).toString("base64");

  const url =
    `https://api.github.com/repos/${owner}/${repo}/contents/${filePath}`;

  const response = await fetch(url, {
    method: "PUT",
    headers: {
      ...githubHeaders(),
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: "Atualiza planilha pelo editor web",
      content: base64Content,
      sha,
      branch,
    }),
  });

  if (!response.ok) {
    const error = await response.text();

    return Response.json(
      {
        error: "Não foi possível salvar a planilha no GitHub.",
        detail: error,
      },
      { status: response.status }
    );
  }

  const result = await response.json();

  return Response.json({
    ok: true,
    commit: result.commit?.sha,
  });
}