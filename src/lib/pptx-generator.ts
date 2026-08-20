import JSZip from "jszip";
import * as XLSX from "xlsx";
import templateAsset from "@/assets/template.pptx.asset.json";

export type LogKind = "ok" | "warn" | "err" | undefined;

export type GeneratorFiles = {
  fv?: File;
  fe?: File;
  ac?: File;
  as?: File;
  sm?: File;
  sa?: File;
};

let log: (msg: string, cls?: LogKind) => void = () => {};

const templatePublicUrl = `${import.meta.env.BASE_URL}template.pptx`;
const TEMPLATE_URLS = [templatePublicUrl, templateAsset.url];

async function loadTemplateBuffer(): Promise<ArrayBuffer> {
  let lastStatus = 0;
  const attemptedUrls: string[] = [];

  for (const url of TEMPLATE_URLS) {
    try {
      const resolvedUrl = new URL(url, window.location.href).href;
      attemptedUrls.push(resolvedUrl);

      const response = await fetch(resolvedUrl, {
        cache: "no-store",
      });

      if (response.ok) {
        return await response.arrayBuffer();
      }

      lastStatus = response.status;
    } catch {
      lastStatus = 0;
    }
  }

  throw new Error(
    `Não foi possível carregar o template. Status: ${lastStatus}. URLs testadas: ${attemptedUrls.join(", ")}.`,
  );
}

const MESES_PT = [
  "janeiro",
  "fevereiro",
  "março",
  "abril",
  "maio",
  "junho",
  "julho",
  "agosto",
  "setembro",
  "outubro",
  "novembro",
  "dezembro",
];

const cap = (v: number) =>
  Number(v || 0).toFixed(2).replace(".", ",");

const fmtPctBR = (v: number) =>
  (Number(v || 0) * 100).toFixed(2).replace(".", ",");

const fmtNumBR = (v: number) =>
  Number(v || 0).toFixed(2).replace(".", ",");

const pctInt = (v: number) =>
  (Number(v || 0) * 100).toFixed(2).replace(".", ",");

/* =========================================================
   DETECÇÃO DE LINHAS/CÉLULAS RISCADAS
========================================================= */

async function detectStruckRows(buf: ArrayBuffer) {
  const out: Record<string, Set<number>> = {};
  const outCells: Record<string, Set<string>> = {};

  const strikeRe = /<strike\b(?![^>]*val="(?:0|false)")/;

  try {
    const zip = await JSZip.loadAsync(buf);

    const stylesFile = zip.file("xl/styles.xml");
    const styles = stylesFile
      ? await stylesFile.async("string")
      : "";

    const struckFonts = new Set<number>();

    const fontsBlock =
      (styles.match(/<fonts[\s\S]*?<\/fonts>/) || [""])[0];

    (
      fontsBlock.match(
        /<font\b[\s\S]*?(?:\/>|<\/font>)/g,
      ) || []
    ).forEach((f, i) => {
      if (strikeRe.test(f)) {
        struckFonts.add(i);
      }
    });

    const struckXfs = new Set<number>();

    const xfsBlock =
      (styles.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [""])[0];

    (
      xfsBlock.match(
        /<xf\b[\s\S]*?(?:\/>|<\/xf>)/g,
      ) || []
    ).forEach((x, i) => {
      const m = x.match(/fontId="(\d+)"/);

      if (m && struckFonts.has(Number(m[1]))) {
        struckXfs.add(i);
      }
    });

    const struckSst = new Set<number>();

    const sstFile = zip.file("xl/sharedStrings.xml");

    if (sstFile) {
      const sst = await sstFile.async("string");

      const items =
        sst.match(
          /<si\b[\s\S]*?(?:\/>|<\/si>)/g,
        ) || [];

      items.forEach((si, i) => {
        if (strikeRe.test(si)) {
          struckSst.add(i);
        }
      });
    }

    if (!struckXfs.size && !struckSst.size) {
      return {
        rows: out,
        cells: outCells,
      };
    }

    const workbookFile = zip.file("xl/workbook.xml");
    const relsFile = zip.file(
      "xl/_rels/workbook.xml.rels",
    );

    if (!workbookFile || !relsFile) {
      return {
        rows: out,
        cells: outCells,
      };
    }

    const wbXml = await workbookFile.async("string");
    const relsXml = await relsFile.async("string");

    const rels: Record<string, string> = {};

    (
      relsXml.match(
        /<Relationship\b[^>]*?(?:\/>|<\/Relationship>)/g,
      ) || []
    ).forEach((r) => {
      const id =
        (r.match(/Id="([^"]+)"/) || [])[1];

      const tgt =
        (r.match(/Target="([^"]+)"/) || [])[1];

      if (id && tgt) {
        rels[id] = tgt
          .replace(/^\/?xl\//, "")
          .replace(/^\.\//, "");
      }
    });

    const sheetTags =
      wbXml.match(
        /<sheet\b[^>]*?(?:\/>|<\/sheet>)/g,
      ) || [];

    for (const st of sheetTags) {
      const name =
        (st.match(/name="([^"]*)"/) || [])[1];

      const rid =
        (st.match(/r:id="([^"]+)"/) || [])[1];

      const path =
        rid && rels[rid]
          ? "xl/" + rels[rid]
          : null;

      if (!name || !path || !zip.file(path)) {
        continue;
      }

      const xml = await zip
        .file(path)!
        .async("string");

      const struckCols = new Set<number>();

      (
        xml.match(/<col\b[^>]*\/>/g) || []
      ).forEach((c) => {
        const sty =
          (c.match(/\bstyle="(\d+)"/) || [])[1];

        if (
          sty == null ||
          !struckXfs.has(Number(sty))
        ) {
          return;
        }

        const min = Number(
          (c.match(/\bmin="(\d+)"/) || [])[1] || 0,
        );

        const max = Number(
          (c.match(/\bmax="(\d+)"/) || [])[1] || min,
        );

        for (
          let k = min;
          k <= max && k - min < 16384;
          k++
        ) {
          struckCols.add(k);
        }
      });

      const set = new Set<number>();
      const cellSet = new Set<string>();

      const rowRe =
        /<row\b([^>]*)(\/>|>([\s\S]*?)<\/row>)/g;

      let m: RegExpExecArray | null;

      while ((m = rowRe.exec(xml)) !== null) {
        const attrs = m[1] || "";
        const body = m[3] || "";

        const rNum = Number(
          (attrs.match(/\br="(\d+)"/) || [])[1],
        );

        if (!rNum) continue;

        let rowStruck = false;

        const rs = attrs.match(/\bs="(\d+)"/);

        if (
          rs &&
          /customFormat="(?:1|true)"/.test(attrs) &&
          struckXfs.has(Number(rs[1]))
        ) {
          rowStruck = true;
        }

        let content = 0;
        let struckContent = 0;

        if (body) {
          const cells =
            body.match(
              /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,
            ) || [];

          for (const c of cells) {
            const ref =
              (c.match(
                /\br="([A-Z]+)\d+"/,
              ) || [])[1];

            const hasVal =
              /<v>|<is>|<t>/.test(c);

            const cs =
              (c.match(
                /\bs="(\d+)"/,
              ) || [])[1];

            let cStruck = false;

            if (
              cs != null &&
              struckXfs.has(Number(cs))
            ) {
              cStruck = true;
            }

            if (
              !cStruck &&
              cs == null &&
              ref &&
              struckCols.has(colIdx(ref))
            ) {
              cStruck = true;
            }

            if (
              !cStruck &&
              /\bt="s"/.test(c)
            ) {
              const v =
                (c.match(
                  /<v>(\d+)<\/v>/,
                ) || [])[1];

              if (
                v != null &&
                struckSst.has(Number(v))
              ) {
                cStruck = true;
              }
            }

            if (
              !cStruck &&
              /\bt="inlineStr"/.test(c) &&
              strikeRe.test(c)
            ) {
              cStruck = true;
            }

            if (hasVal) {
              content++;
            }

            if (cStruck) {
              if (hasVal) {
                struckContent++;
              }

              if (ref) {
                cellSet.add(
                  (rNum - 1) +
                    "," +
                    (colIdx(ref) - 1),
                );
              }
            }
          }
        }

        if (
          rowStruck ||
          (content > 0 &&
            struckContent === content)
        ) {
          set.add(rNum - 1);
        }
      }

      if (set.size) {
        out[name] = set;
      }

      if (cellSet.size) {
        outCells[name] = cellSet;
      }
    }
  } catch {
    // silencioso
  }

  return {
    rows: out,
    cells: outCells,
  };
}

function colIdx(letters: string) {
  let n = 0;

  for (let i = 0; i < letters.length; i++) {
    n =
      n * 26 +
      (letters.charCodeAt(i) - 64);
  }

  return n;
}

function isMarkdownStruck(value: unknown) {
  if (value == null) return false;

  const text = String(value).trim();

  return (
    /^~~[\s\S]+~~$/.test(text) &&
    text.slice(2, -2).length > 0
  );
}

function mergeMarkdownStruck(
  wb: XLSX.WorkBook,
  rowsBySheet: Record<string, Set<number>>,
  cellsBySheet: Record<string, Set<string>>,
) {
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];

    if (!ws || !ws["!ref"]) continue;

    const range = XLSX.utils.decode_range(
      ws["!ref"],
    );

    const grid = XLSX.utils.sheet_to_json(
      ws,
      {
        header: 1,
        defval: null,
        raw: false,
      },
    ) as unknown[][];

    const rowSet =
      rowsBySheet[name] || new Set<number>();

    const cellSet =
      cellsBySheet[name] ||
      new Set<string>();

    for (let i = 0; i < grid.length; i++) {
      const row = grid[i] || [];

      let filled = 0;
      let markdownStruck = 0;

      for (let j = 0; j < row.length; j++) {
        const value = row[j];

        if (
          value == null ||
          String(value).trim() === ""
        ) {
          continue;
        }

        filled++;

        if (isMarkdownStruck(value)) {
          markdownStruck++;

          cellSet.add(
            range.s.r +
              i +
              "," +
              (range.s.c + j),
          );
        }
      }

      if (
        filled > 0 &&
        markdownStruck === filled
      ) {
        rowSet.add(range.s.r + i);
      }
    }

    if (rowSet.size) {
      rowsBySheet[name] = rowSet;
    }

    if (cellSet.size) {
      cellsBySheet[name] = cellSet;
    }
  }
}

function struckSet(
  wb: XLSX.WorkBook,
  sheetName: string,
) {
  return (
    (wb as any).__struck?.[sheetName] ||
    null
  );
}

function sheetBase(ws: XLSX.WorkSheet) {
  try {
    return XLSX.utils.decode_range(
      ws["!ref"]!,
    ).s.r;
  } catch {
    return 0;
  }
}

function totalStruck(wb: XLSX.WorkBook) {
  let n = 0;

  const struck = (wb as any).__struck;

  if (struck) {
    for (const k in struck) {
      n += struck[k].size;
    }
  }

  return n;
}

/* =========================================================
   LEITURA XLSX
========================================================= */

async function readXlsx(
  file: File | undefined,
  nome?: string,
) {
  if (!file) {
    throw new Error(
      `Arquivo não enviado: ${
        nome ?? "desconhecido"
      }`,
    );
  }

  if (!(file instanceof File)) {
    throw new Error(
      `O arquivo "${
        nome ?? "desconhecido"
      }" não é um File válido.`,
    );
  }

  const buf = await file.arrayBuffer();

  const wb = XLSX.read(buf, {
    type: "array",
    cellDates: true,
  });

  const det = await detectStruckRows(buf);

  mergeMarkdownStruck(
    wb,
    det.rows,
    det.cells,
  );

  (wb as any).__struck = det.rows;
  (wb as any).__struckCells = det.cells;

  for (const nm in det.cells) {
    const ws = wb.Sheets[nm];

    if (!ws) continue;

    det.cells[nm].forEach((k) => {
      const p = k.split(",");

      const addr =
        XLSX.utils.encode_cell({
          r: Number(p[0]),
          c: Number(p[1]),
        });

      if (ws[addr]) {
        delete ws[addr];
      }
    });
  }

  return wb;
}

async function readOptionalXlsx(
  file: File | undefined,
  nome: string,
) {
  if (!file) {
    log(
      `Arquivo "${nome}" não enviado. Conteúdo original do template será mantido.`,
      "warn",
    );

    return null;
  }

  try {
    return await readXlsx(file, nome);
  } catch (error) {
    log(
      `Erro ao ler "${nome}": ${
        error instanceof Error
          ? error.message
          : String(error)
      }`,
      "err",
    );

    return null;
  }
}

/* =========================================================
   SCORE MENSAL
========================================================= */

function parseScoreMensal(
  wb: XLSX.WorkBook,
) {
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];

  const struck = struckSet(wb, name);
  const base = sheetBase(ws);

  const grid = XLSX.utils.sheet_to_json(
    ws,
    {
      header: 1,
      defval: null,
      raw: true,
    },
  ) as unknown[][];

  const rows: {
    date: Date;
    score: number;
  }[] = [];

  for (let i = 0; i < grid.length; i++) {
    if (
      struck &&
      struck.has(base + i)
    ) {
      continue;
    }

    const r = grid[i];

    if (
      r &&
      r[0] instanceof Date &&
      typeof r[1] === "number"
    ) {
      rows.push({
        date: r[0] as Date,
        score: r[1] as number,
      });
    }
  }

  rows.sort(
    (a, b) =>
      a.date.getTime() -
      b.date.getTime(),
  );

  rows.pop();

  return rows;
}

/* =========================================================
   SCORE FECHADO
========================================================= */

function parseScoreFechado(
  wb: XLSX.WorkBook,
) {
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];

  const grid = XLSX.utils.sheet_to_json(
    ws,
    {
      header: 1,
      defval: null,
      raw: true,
    },
  ) as unknown[][];

  let headerRow = -1;

  for (let i = 0; i < grid.length; i++) {
    if (grid[i]?.[1] === "Score") {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    throw new Error(
      "Cabeçalho não encontrado.",
    );
  }

  const r = grid[headerRow + 1];

  if (!r) {
    throw new Error(
      "Linha de dados não encontrada.",
    );
  }

  return {
    industria: r[0] ?? null,
    score: r[1],
    frequencia: r[2],
    acuracidade: r[3],
    consistencia: r[4],
    validacao: r[5],
  };
}

/* =========================================================
   VALIDAÇÃO
========================================================= */

function parseValidacao(
  wb: XLSX.WorkBook,
) {
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];

  const struck = struckSet(wb, name);
  const base = sheetBase(ws);

  const isStruck = (i: number) =>
    !!(
      struck &&
      struck.has(base + i)
    );

  const grid = XLSX.utils.sheet_to_json(
    ws,
    {
      header: 1,
      defval: null,
      raw: false,
    },
  ) as unknown[][];

  let latest: string | null = null;

  for (let i = 1; i < grid.length; i++) {
    if (isStruck(i)) continue;

    const m = grid[i]?.[0];

    if (
      typeof m === "string" &&
      /^\d{2}-\d{2}$/.test(m.trim())
    ) {
      const v = m.trim();

      if (!latest || v > latest) {
        latest = v;
      }
    }
  }

  let dentro = 0;
  let fora = 0;
  let pend = 0;
  let div = 0;
  let total = 0;

  const tabela: {
    cnpj: string;
    nome: string;
    status: string;
  }[] = [];

  let riscadas = 0;

  for (let i = 1; i < grid.length; i++) {
    if (isStruck(i)) {
      riscadas++;
      continue;
    }

    const r = grid[i];

    if (!r) continue;

    const mes =
      typeof r[0] === "string"
        ? r[0].trim()
        : r[0];

    if (mes !== latest) continue;

    const nome =
      r[2] != null
        ? String(r[2]).trim()
        : "";

    const cnpj =
      r[3] != null
        ? String(r[3]).trim()
        : "";

    const st =
      r[4] != null
        ? String(r[4]).trim()
        : "";

    if (!cnpj && !nome) continue;

    total++;

    const low = st.toLowerCase();

    if (low.includes("diverg")) {
      div++;

      tabela.push({
        cnpj,
        nome,
        status: "Divergência",
      });
    } else if (low.includes("pend")) {
      pend++;

      tabela.push({
        cnpj,
        nome,
        status: "Pendente",
      });
    } else if (low.includes("fora")) {
      fora++;
    } else if (
      low.includes("dentro") ||
      low.includes("valid")
    ) {
      dentro++;
    }
  }

  tabela.sort(
    (a, b) =>
      a.status === b.status
        ? a.nome.localeCompare(b.nome)
        : a.status.localeCompare(b.status),
  );

  const d = total || 1;

  return {
    latest,
    total,
    riscadas,
    dentro,
    fora,
    pend,
    div,
    tabela,
    pctDentro: dentro / d,
    pctDentro: dentro / d,
pctFora: fora / d,
pctPend: pend / d,
pctDiv: div / d,
pctValidado: (dentro + fora) / d,
  };
}

/* =========================================================
   FREQUÊNCIA
========================================================= */

function parseFrequenciaDias(
  wb: XLSX.WorkBook,
) {
  const name = wb.SheetNames[0];
  const ws = wb.Sheets[name];

  const struck = struckSet(wb, name);
  const base = sheetBase(ws);

  const grid = XLSX.utils.sheet_to_json(
    ws,
    {
      header: 1,
      defval: null,
      raw: false,
    },
  ) as unknown[][];

  let mesLabel: string | null = null;

  const filt = grid[0]?.[0];

  if (typeof filt === "string") {
    const m = filt.match(
      /Mês\s*=\s*([A-Za-zç]+\/\d{4})/,
    );

    if (m) {
      mesLabel = m[1];
    }
  }

  let colFreq = -1;

  for (
    let i = 0;
    i < Math.min(6, grid.length) &&
    colFreq < 0;
    i++
  ) {
    const r = grid[i];

    if (!r) continue;

    for (
      let j = 0;
      j < Math.min(r.length, 6);
      j++
    ) {
      const v = r[j];

      if (v == null) continue;

      const t = String(v).toLowerCase();

      if (
        t.includes("frequ") ||
        t.includes("%")
      ) {
        colFreq = j;
        break;
      }
    }
  }

  const parsePct = (
    v: unknown,
  ): number | null => {
    if (v == null) return null;

    if (typeof v === "number") {
      return v > 1 ? v / 100 : v;
    }

    let t = String(v).trim();

    if (
      !t ||
      isMarkdownStruck(t)
    ) {
      return null;
    }

    const hasPct = t.includes("%");

    t = t
      .replace("%", "")
      .replace(/\s/g, "")
      .replace(/\./g, "")
      .replace(",", ".");

    const n = parseFloat(t);

    if (isNaN(n)) return null;

    return hasPct || n > 1
      ? n / 100
      : n;
  };

  const byCnpj = new Map<
    string,
    {
      cnpj: string;
      nome: string;
      ausentes: number;
      considerados: number;
      pctPlan: number | null;
    }
  >();

  const isStruck = (i: number) =>
    !!(
      struck &&
      struck.has(base + i)
    );

  for (
    let i = 5;
    i < grid.length;
    i++
  ) {
    if (isStruck(i)) continue;

    const r = grid[i];

    if (!r) continue;

    let filled = 0;
    let struckCells = 0;

    for (let j = 0; j < r.length; j++) {
      const v = r[j];

      if (
        v == null ||
        String(v).trim() === ""
      ) {
        continue;
      }

      filled++;

      if (isMarkdownStruck(v)) {
        struckCells++;
      }
    }

    if (
      filled > 0 &&
      struckCells === filled
    ) {
      continue;
    }

    const nome =
      !isMarkdownStruck(r[0]) &&
      r[0] != null
        ? String(r[0]).trim()
        : "";

    const cnpj =
      !isMarkdownStruck(r[1]) &&
      r[1] != null
        ? String(r[1]).trim()
        : "";

    if (!cnpj) continue;

    let ausentes = 0;
    let considerados = 0;

    for (let j = 5; j < r.length; j++) {
      const v = r[j];

      if (v == null) continue;

      if (isMarkdownStruck(v)) {
        continue;
      }

      const s = String(v)
        .trim()
        .toUpperCase();

      if (!s) continue;

      considerados++;

      if (s === "I") {
        ausentes++;
      }
    }

    if (considerados === 0) {
      continue;
    }

    const pctPlan =
      colFreq >= 0
        ? parsePct(r[colFreq])
        : null;

    const prev = byCnpj.get(cnpj);

    if (
      !prev ||
      ausentes > prev.ausentes
    ) {
      byCnpj.set(cnpj, {
        cnpj,
        nome,
        ausentes,
        considerados,
        pctPlan,
      });
    }
  }

  const list = [
    ...byCnpj.values(),
  ];

  let atualizados = 0;
  let atencao = 0;
  let ofensores = 0;
  let somaAus = 0;
  let somaCon = 0;

  for (const d of list) {
    if (d.ausentes >= 5) {
      ofensores++;
    } else if (d.ausentes >= 1) {
      atencao++;
    } else {
      atualizados++;
    }

    somaAus += d.ausentes;
    somaCon += d.considerados;
  }

  const pctFreqRow = (d: {
    pctPlan: number | null;
    considerados: number;
    ausentes: number;
  }) =>
    d.pctPlan != null
      ? Number(d.pctPlan)
      : d.considerados
        ? (d.considerados -
            d.ausentes) /
          d.considerados
        : 0;

  const pctAus = (d: {
    pctPlan: number | null;
    considerados: number;
    ausentes: number;
  }) =>
    1 - pctFreqRow(d);

  const ofensoresList = list
    .filter(
      (d) => d.ausentes >= 5,
    )
    .sort(
      (a, b) =>
        pctAus(b) - pctAus(a) ||
        b.ausentes - a.ausentes ||
        a.nome.localeCompare(b.nome),
    );

  const pctFreq = somaCon
    ? (somaCon - somaAus) /
      somaCon
    : 0;

  return {
    mesLabel,
    list,
    atualizados,
    atencao,
    ofensores,
    total: list.length,
    ofensoresList,
    pctFreq,
  };
}

/* =========================================================
   ASSOCIAÇÕES
========================================================= */

function parseAssoc(
  wb: XLSX.WorkBook,
) {
  const resumoName =
    wb.Sheets["Resumo"]
      ? "Resumo"
      : wb.SheetNames[0];

  const resumoSheet =
    wb.Sheets[resumoName];

  const resumoStruck =
    struckSet(wb, resumoName);

  const resumoBase =
    sheetBase(resumoSheet);

  const resumoGrid =
    XLSX.utils.sheet_to_json(
      resumoSheet,
      {
        header: 1,
        defval: null,
      },
    ) as unknown[][];

  let dist = 0;
  let ind = 0;

  for (
    let ri = 0;
    ri < resumoGrid.length;
    ri++
  ) {
    if (
      resumoStruck &&
      resumoStruck.has(
        resumoBase + ri,
      )
    ) {
      continue;
    }

    const r = resumoGrid[ri];

    if (!r) continue;

    const nums = r.filter(
      (x) => typeof x === "number",
    ) as number[];

    if (nums.length >= 3) {
      ind = Math.round(nums[1]);
      dist = Math.round(nums[2]);
      break;
    }
  }

  const itemsName =
    wb.Sheets["Associação de Itens"]
      ? "Associação de Itens"
      : wb.SheetNames[1];

  const itemsSheet =
    itemsName
      ? wb.Sheets[itemsName]
      : undefined;

  const itemsStruck =
    itemsSheet
      ? struckSet(wb, itemsName)
      : null;

  const itemsBase =
    itemsSheet
      ? sheetBase(itemsSheet)
      : 0;

  let sim = 0;
  let nao = 0;
  let rejeit = 0;

  if (itemsSheet) {
    const grid =
      XLSX.utils.sheet_to_json(
        itemsSheet,
        {
          header: 1,
          defval: null,
        },
      ) as unknown[][];

    let hdr = -1;
    let cAssoc = -1;

    for (
      let i = 0;
      i < grid.length;
      i++
    ) {
      const r = grid[i];

      if (!r) continue;

      const idx = r.findIndex(
        (x) =>
          typeof x === "string" &&
          x.trim() ===
            "Item foi Associado?",
      );

      if (idx >= 0) {
        hdr = i;
        cAssoc = idx;
        break;
      }
    }

    if (hdr >= 0) {
      for (
        let i = hdr + 1;
        i < grid.length;
        i++
      ) {
        if (
          itemsStruck &&
          itemsStruck.has(
            itemsBase + i,
          )
        ) {
          continue;
        }

        const r = grid[i];

        if (!r) continue;

        const v = r[cAssoc];

        if (
          v == null ||
          v === ""
        ) {
          if (
            r.some(
              (x) =>
                x != null &&
                x !== "",
            )
          ) {
            rejeit++;
          }
        } else {
          const s = String(v)
            .trim()
            .toUpperCase();

          if (s === "SIM") {
            sim++;
          } else if (
            s === "NÃO" ||
            s === "NAO"
          ) {
            nao++;
          }
        }
      }
    }
  }

  return {
    distribuidores: dist,
    industrias: ind,
    associados: sim,
    naoIdent: nao,
    rejeitados: rejeit,
    impactados: sim,
  };
}

/* =========================================================
   XML
========================================================= */

function xmlEscape(s: unknown) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function replaceRun(
  xml: string,
  oldText: string,
  newText: unknown,
) {
  const target =
    "<a:t>" +
    oldText +
    "</a:t>";

  const rep =
    "<a:t>" +
    xmlEscape(newText) +
    "</a:t>";

  const i = xml.indexOf(target);

  if (i < 0) {
    log(
      "  ! não encontrado: " +
        oldText,
      "warn",
    );

    return xml;
  }

  return (
    xml.slice(0, i) +
    rep +
    xml.slice(i + target.length)
  );
}

function replaceLastRun(
  xml: string,
  oldText: string,
  newText: unknown,
) {
  const target =
    "<a:t>" +
    oldText +
    "</a:t>";

  const i =
    xml.lastIndexOf(target);

  if (i < 0) {
    log(
      "  ! não encontrado (último): " +
        oldText,
      "warn",
    );

    return xml;
  }

  return (
    xml.slice(0, i) +
    "<a:t>" +
    xmlEscape(newText) +
    "</a:t>" +
    xml.slice(i + target.length)
  );
}

function replaceRunsSeq(
  xml: string,
  oldText: string,
  values: unknown[],
) {
  let out = xml;
  let cursor = 0;
  let from = 0;

  const target =
    "<a:t>" +
    oldText +
    "</a:t>";

  while (cursor < values.length) {
    const i = out.indexOf(
      target,
      from,
    );

    if (i < 0) {
      log(
        "  ! run não encontrado: " +
          oldText,
        "warn",
      );

      break;
    }

    const rep =
      "<a:t>" +
      xmlEscape(values[cursor]) +
      "</a:t>";

    out =
      out.slice(0, i) +
      rep +
      out.slice(
        i + target.length,
      );

    from = i + rep.length;
    cursor++;
  }

  return out;
}

function replaceNumberBeforeLabel(
  xml: string,
  label: string,
  newVal: unknown,
) {
  const labelTarget =
    "<a:t>" +
    label +
    "</a:t>";

  const labelIdx =
    xml.indexOf(labelTarget);

  if (labelIdx < 0) {
    log(
      "  ! rótulo não encontrado: " +
        label,
      "warn",
    );

    return xml;
  }

  const before =
    xml.slice(0, labelIdx);

  const openIdx =
    before.lastIndexOf("<a:t>");

  const closeIdx =
    before.indexOf(
      "</a:t>",
      openIdx,
    );

  if (
    openIdx < 0 ||
    closeIdx < 0
  ) {
    log(
      "  ! número não encontrado antes de: " +
        label,
      "warn",
    );

    return xml;
  }

  const prefix = before.slice(
    openIdx + 5,
    closeIdx,
  );

  const leadSpace =
    (prefix.match(/^\s*/) || [
      "",
    ])[0];

  const newRun =
    "<a:t>" +
    leadSpace +
    xmlEscape(String(newVal)) +
    "</a:t>";

  return (
    xml.slice(0, openIdx) +
    newRun +
    xml.slice(closeIdx + 6)
  );
}

function replaceNumberAfterLabel(
  xml: string,
  label: string,
  newVal: unknown,
) {
  const labelTarget =
    "<a:t>" +
    label +
    "</a:t>";

  const labelIdx =
    xml.indexOf(labelTarget);

  if (labelIdx < 0) {
    log(
      "  ! rótulo não encontrado: " +
        label,
      "warn",
    );

    return xml;
  }

  const after =
    labelIdx + labelTarget.length;

  const openIdx =
    xml.indexOf("<a:t>", after);

  const closeIdx =
    xml.indexOf("</a:t>", openIdx);

  if (
    openIdx < 0 ||
    closeIdx < 0
  ) {
    log(
      "  ! número não encontrado após: " +
        label,
      "warn",
    );

    return xml;
  }

  return (
    xml.slice(0, openIdx) +
    "<a:t>" +
    xmlEscape(String(newVal)) +
    "</a:t>" +
    xml.slice(closeIdx + 6)
  );
}

function fillMesAno(
  xml: string,
  mesCap: string,
  ano: string,
) {
  xml = replaceRun(
    xml,
    "mes",
    mesCap,
  );

  xml = replaceRun(
    xml,
    " ano",
    " " + ano,
  );

  return xml;
}

/* =========================================================
   TABELAS
========================================================= */

function countTableDataRows(
  xml: string,
) {
  const tblStart =
    xml.indexOf("<a:tbl>");

  if (tblStart < 0) return 0;

  const tblEnd =
    xml.indexOf(
      "</a:tbl>",
      tblStart,
    ) + 8;

  const tbl = xml.slice(
    tblStart,
    tblEnd,
  );

  const trs =
    tbl.match(
      /<a:tr[ >]/g,
    ) || [];

  return Math.max(
    0,
    trs.length - 1,
  );
}

function fillTable(
  xml: string,
  rows: unknown[][],
) {
  const tblStart =
    xml.indexOf("<a:tbl>");

  if (tblStart < 0) {
    log(
      "  ! tabela não encontrada",
      "warn",
    );

    return xml;
  }

  const tblEnd =
    xml.indexOf(
      "</a:tbl>",
      tblStart,
    ) + 8;

  let tbl = xml.slice(
    tblStart,
    tblEnd,
  );

  const trRe =
    /<a:tr[ >][\s\S]*?<\/a:tr>/g;

  let dataIdx = -1;

  tbl = tbl.replace(
    trRe,
    (tr) => {
      dataIdx++;

      if (dataIdx === 0) {
        return tr;
      }

      const data =
        rows[dataIdx - 1];

      if (
        !data ||
        !data.some(
          (v) =>
            v != null &&
            String(v).trim() !== "",
        )
      ) {
        return "";
      }

      let cellIdx = -1;

      return tr.replace(
        /<a:tc>[\s\S]*?<\/a:tc>/g,
        (tc) => {
          cellIdx++;

          const val =
            data &&
            data[cellIdx] != null
              ? String(
                  data[cellIdx],
                )
              : " ";

          if (!/<a:t>/.test(tc)) {
            return tc.replace(
              /<a:endParaRPr([^>]*)(\/>|>[\s\S]*?<\/a:endParaRPr>)/,
              (
                m,
                attrs,
                rest,
              ) => {
                const inner =
                  rest === "/>"
                    ? ""
                    : rest.slice(
                        1,
                        rest.lastIndexOf(
                          "</a:endParaRPr>",
                        ),
                      );

                const rPr =
                  "<a:rPr" +
                  attrs +
                  (inner
                    ? ">" +
                      inner +
                      "</a:rPr>"
                    : "/>");

                return (
                  "<a:r>" +
                  rPr +
                  "<a:t>" +
                  xmlEscape(val) +
                  "</a:t></a:r>" +
                  m
                );
              },
            );
          }

          let done = false;

          return tc.replace(
            /<a:t>[\s\S]*?<\/a:t>/g,
            () => {
              if (done) {
                return "<a:t></a:t>";
              }

              done = true;

              return (
                "<a:t>" +
                xmlEscape(val) +
                "</a:t>"
              );
            },
          );
        },
      );
    },
  );

  return (
    xml.slice(0, tblStart) +
    tbl +
    xml.slice(tblEnd)
  );
}

function removeTableFromSlide(xml: string) {
  /*
   * A tabela do PowerPoint normalmente fica dentro de:
   *
   * <p:graphicFrame>
   *   ...
   *   <a:tbl>
   *     ...
   *   </a:tbl>
   * </p:graphicFrame>
   *
   * Quando não há dados, removemos o graphicFrame inteiro,
   * evitando deixar uma tabela vazia no slide.
   */

  const tblStart = xml.indexOf("<a:tbl>");

  if (tblStart < 0) {
    log(
      "  ! tabela de validação não encontrada no slide.",
      "warn",
    );

    return xml;
  }

  /*
   * Procura o <p:graphicFrame> que contém a tabela.
   */
  const frameStart = xml.lastIndexOf(
    "<p:graphicFrame",
    tblStart,
  );

  const frameEndStart = xml.indexOf(
    "</p:graphicFrame>",
    tblStart,
  );

  if (
    frameStart >= 0 &&
    frameEndStart >= 0
  ) {
    const frameEnd =
      frameEndStart +
      "</p:graphicFrame>".length;

    return (
      xml.slice(0, frameStart) +
      xml.slice(frameEnd)
    );
  }

  /*
   * Fallback:
   * se por algum motivo a tabela não estiver
   * dentro de um graphicFrame, remove somente
   * a tabela.
   */
  const tblEndStart = xml.indexOf(
    "</a:tbl>",
    tblStart,
  );

  if (tblEndStart >= 0) {
    const tblEnd =
      tblEndStart +
      "</a:tbl>".length;

    return (
      xml.slice(0, tblStart) +
      xml.slice(tblEnd)
    );
  }

  return xml;
}

async function fillValidacaoSlide(
  zip: JSZip,
  slideNum: number,
  rows: unknown[][],
) {
  const file = zip.file(
    `ppt/slides/slide${slideNum}.xml`,
  );

  if (!file) {
    log(
      `  ! slide ${slideNum} não encontrado`,
      "warn",
    );
    return;
  }

  let xml = await file.async("string");

  /*
   * Se não existem dados de validação,
   * remove todas as tabelas do slide.
   */
  if (!rows || rows.length === 0) {
    let removedCount = 0;

    const tblRe =
      /<p:graphicFrame[\s\S]*?<a:tbl>[\s\S]*?<\/a:tbl>[\s\S]*?<\/p:graphicFrame>/g;

    xml = xml.replace(tblRe, () => {
      removedCount++;
      return "";
    });

    zip.file(
      `ppt/slides/slide${slideNum}.xml`,
      xml,
    );

    log(
      "  Validação: nenhum registro encontrado. Tabelas removidas.",
      "warn",
    );

    return;
  }

  /*
   * ============================================================
   * LOCALIZA TODAS AS TABELAS DO SLIDE
   * ============================================================
   */

  const tableBlocks: {
    start: number;
    end: number;
    xml: string;
  }[] = [];

  const tableRe =
    /<a:tbl>[\s\S]*?<\/a:tbl>/g;

  let match: RegExpExecArray | null;

  while ((match = tableRe.exec(xml)) !== null) {
    tableBlocks.push({
      start: match.index,
      end: match.index + match[0].length,
      xml: match[0],
    });
  }

  if (!tableBlocks.length) {
    log(
      "  ! nenhuma tabela encontrada no slide de validação.",
      "warn",
    );

    return;
  }

  /*
   * ============================================================
   * DESCOBRE QUANTAS LINHAS DE DADOS CADA TABELA POSSUI
   * ============================================================
   *
   * A primeira linha é o cabeçalho.
   */
  const capacities = tableBlocks.map((table) =>
    countTableDataRows(table.xml),
  );

  /*
   * Caso alguma tabela não tenha sido reconhecida,
   * usa 29 como fallback.
   */
  const normalizedCapacities = capacities.map(
    (capacity) =>
      capacity > 0 ? capacity : 29,
  );

  /*
   * ============================================================
   * DISTRIBUI OS DADOS ENTRE AS TABELAS
   * ============================================================
   */

  let rowOffset = 0;

  const replacements: {
    start: number;
    end: number;
    replacement: string;
  }[] = [];

  for (
    let i = 0;
    i < tableBlocks.length;
    i++
  ) {
    const table = tableBlocks[i];

    const capacity =
      normalizedCapacities[i];

    const tableRows = rows.slice(
      rowOffset,
      rowOffset + capacity,
    );

    rowOffset += tableRows.length;

    /*
     * ==========================================================
     * TABELA SEM DADOS
     * ==========================================================
     *
     * Remove o graphicFrame inteiro.
     */
    if (tableRows.length === 0) {
      const frameStart = xml.lastIndexOf(
        "<p:graphicFrame",
        table.start,
      );

      const frameEndStart = xml.indexOf(
        "</p:graphicFrame>",
        table.end,
      );

      if (
        frameStart >= 0 &&
        frameEndStart >= 0
      ) {
        const frameEnd =
          frameEndStart +
          "</p:graphicFrame>".length;

        replacements.push({
          start: frameStart,
          end: frameEnd,
          replacement: "",
        });
      } else {
        replacements.push({
          start: table.start,
          end: table.end,
          replacement: "",
        });
      }

      continue;
    }

    /*
     * ==========================================================
     * TABELA COM DADOS
     * ==========================================================
     */

    const filled = fillTable(
      table.xml,
      tableRows,
    );

    replacements.push({
      start: table.start,
      end: table.end,
      replacement: filled,
    });
  }

  /*
   * ============================================================
   * APLICA AS ALTERAÇÕES DE TRÁS PARA FRENTE
   * ============================================================
   *
   * Isso evita alterar os índices das posições anteriores.
   */
  replacements
    .sort((a, b) => b.start - a.start)
    .forEach((r) => {
      xml =
        xml.slice(0, r.start) +
        r.replacement +
        xml.slice(r.end);
    });

  zip.file(
    `ppt/slides/slide${slideNum}.xml`,
    xml,
  );

  
}
/* =========================================================
   SLIDES
========================================================= */

async function addSlideCopy(
  zip: JSZip,
  srcNum: number,
  newXml: string,
  afterNum?: number,
) {
  let n = 1;

  while (
    zip.file(
      "ppt/slides/slide" +
        n +
        ".xml",
    )
  ) {
    n++;
  }

  const newNum = n;

  zip.file(
    "ppt/slides/slide" +
      newNum +
      ".xml",
    newXml,
  );

  const relsFile = zip.file(
    "ppt/slides/_rels/slide" +
      srcNum +
      ".xml.rels",
  );

  if (relsFile) {
    let srcRels =
      await relsFile.async(
        "string",
      );

    srcRels = srcRels.replace(
      /<Relationship[^>]*Target="\.\.\/(notesSlides|comments)\/[^"]*"[^>]*\/>/g,
      "",
    );

    zip.file(
      "ppt/slides/_rels/slide" +
        newNum +
        ".xml.rels",
      srcRels,
    );
  }

  const ctFile = zip.file(
    "[Content_Types].xml",
  );

  if (ctFile) {
    let ct =
      await ctFile.async(
        "string",
      );

    ct = ct.replace(
      "</Types>",
      '<Override PartName="/ppt/slides/slide' +
        newNum +
        '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    );

    zip.file(
      "[Content_Types].xml",
      ct,
    );
  }

  const prelsFile = zip.file(
    "ppt/_rels/presentation.xml.rels",
  );

  if (!prelsFile) {
    return newNum;
  }

  let prels =
    await prelsFile.async(
      "string",
    );

  const ids = [
    ...prels.matchAll(
      /Id="rId(\d+)"/g,
    ),
  ].map((m) =>
    parseInt(m[1], 10),
  );

  const newRid =
    "rId" +
    ((ids.length
      ? Math.max(...ids)
      : 0) +
      1);

  prels = prels.replace(
    "</Relationships>",
    '<Relationship Id="' +
      newRid +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' +
      newNum +
      '.xml"/></Relationships>',
  );

  const anchorNum =
    afterNum || srcNum;

  const srcRelMatch =
    prels.match(
      new RegExp(
        'Id="(rId\\d+)"[^>]*Target="slides/slide' +
          anchorNum +
          '\\.xml"',
      ),
    );

  zip.file(
    "ppt/_rels/presentation.xml.rels",
    prels,
  );

  const presFile = zip.file(
    "ppt/presentation.xml",
  );

  if (!presFile) {
    return newNum;
  }

  let pres =
    await presFile.async(
      "string",
    );

  const sldIds = [
    ...pres.matchAll(
      /<p:sldId id="\d+" r:id="rId\d+"\/>/g,
    ),
  ].map((m) => m[0]);

  const idMatches = [
    ...pres.matchAll(
      /<p:sldId id="(\d+)"/g,
    ),
  ];

  const maxId = idMatches.length
    ? Math.max(
        ...idMatches.map(
          (m) =>
            parseInt(
              m[1],
              10,
            ),
        ),
      )
    : 0;

  const newSldId =
    '<p:sldId id="' +
    (maxId + 1) +
    '" r:id="' +
    newRid +
    '"/>';

  let anchor: string | null =
    null;

  if (srcRelMatch) {
    anchor =
      sldIds.find(
        (s) =>
          s.includes(
            'r:id="' +
              srcRelMatch[1] +
              '"',
          ),
      ) || null;
  }

  if (anchor) {
    const idx =
      pres.indexOf(anchor) +
      anchor.length;

    pres =
      pres.slice(0, idx) +
      newSldId +
      pres.slice(idx);
  } else {
    pres = pres.replace(
      "</p:sldIdLst>",
      newSldId +
        "</p:sldIdLst>",
    );
  }

  zip.file(
    "ppt/presentation.xml",
    pres,
  );

  return newNum;
}

function setTableHeaderLast(
  xml: string,
  text: string,
) {
  const tblStart =
    xml.indexOf("<a:tbl>");

  if (tblStart < 0) {
    return xml;
  }

  const tblEnd =
    xml.indexOf(
      "</a:tbl>",
      tblStart,
    ) + 8;

  let tbl = xml.slice(
    tblStart,
    tblEnd,
  );

  const trMatch =
    tbl.match(
      /<a:tr[ >][\s\S]*?<\/a:tr>/,
    );

  if (!trMatch) {
    return xml;
  }

  let tr = trMatch[0];

  const tcs =
    tr.match(
      /<a:tc>[\s\S]*?<\/a:tc>/g,
    );

  if (!tcs || !tcs.length) {
    return xml;
  }

  const lastTc =
    tcs[tcs.length - 1];

  if (!/<a:t>/.test(lastTc)) {
    return xml;
  }

  let done = false;

  const newTc =
    lastTc.replace(
      /<a:t>[\s\S]*?<\/a:t>/g,
      () => {
        if (done) {
          return "<a:t></a:t>";
        }

        done = true;

        return (
          "<a:t>" +
          xmlEscape(text) +
          "</a:t>"
        );
      },
    );

  tr = tr.replace(
    lastTc,
    newTc,
  );

  tbl = tbl.replace(
    trMatch[0],
    tr,
  );

  return (
    xml.slice(0, tblStart) +
    tbl +
    xml.slice(tblEnd)
  );
}

async function fillTableAcrossSlides(
  zip: JSZip,
  slideNums: number[],
  rows: unknown[][],
  label: string,
  prepare?: (
    xml: string,
  ) => string,
) {
  const sources: string[] = [];

  for (const n of slideNums) {
    const file = zip.file(
      "ppt/slides/slide" +
        n +
        ".xml",
    );

    if (!file) continue;

    let s =
      await file.async(
        "string",
      );

    if (prepare) {
      s = prepare(s);
    }

    sources.push(s);
  }

  if (!sources.length) {
    return;
  }

  const perPage =
    countTableDataRows(
      sources[0],
    ) || 29;

  const pages: unknown[][][] =
    [];

  for (
    let i = 0;
    i < rows.length;
    i += perPage
  ) {
    pages.push(
      rows.slice(
        i,
        i + perPage,
      ),
    );
  }

  if (!pages.length) {
    pages.push([]);
  }

  const used = Math.min(
    pages.length,
    slideNums.length,
  );

  for (
    let p = 0;
    p < used;
    p++
  ) {
    zip.file(
      "ppt/slides/slide" +
        slideNums[p] +
        ".xml",
      fillTable(
        sources[p],
        pages[p],
      ),
    );
  }

  let lastAnchor =
    slideNums[
      used - 1
    ];

  const lastSrc =
    sources[
      sources.length - 1
    ];

  for (
    let p = used;
    p < pages.length;
    p++
  ) {
    const num =
      await addSlideCopy(
        zip,
        slideNums[
          slideNums.length - 1
        ],
        fillTable(
          lastSrc,
          pages[p],
        ),
        lastAnchor,
      );

    lastAnchor = num;
  }

 
}


// ============================================================
// Atualiza os pontos de um gráfico do PowerPoint
// ============================================================
function setPointsInSection(
  xml: string,
  sectionStart: string,
  sectionEndTag: string,
  values: any[],
): string {
  const start = xml.indexOf(
    sectionStart,
  );

  if (start < 0) {
    log(
      "  ! seção não encontrada: " +
        sectionStart,
      "warn",
    );
    return xml;
  }

  const end = xml.indexOf(
    sectionEndTag,
    start,
  );

  if (end < 0) {
    log(
      "  ! fim da seção não encontrado: " +
        sectionEndTag,
      "warn",
    );
    return xml;
  }

  let middle = xml.slice(
    start,
    end,
  );

  /*
   * Captura os atributos existentes
   * no primeiro <c:pt>.
   */
  const first =
    middle.match(
      /<c:pt\s+idx="\d+"([^>]*)>/,
    );

  const extraAttrs =
    first ? first[1] : "";

  /*
   * Cria novamente todos os pontos.
   */
  const pts = values
    .map(
      (v, i) =>
        '<c:pt idx="' +
        i +
        '"' +
        extraAttrs +
        "><c:v>" +
        xmlEscape(String(v)) +
        "</c:v></c:pt>",
    )
    .join("");

  /*
   * Remove pontos antigos.
   */
  middle = middle.replace(
    /<c:pt\s+idx="\d+"([^>]*)>\s*<c:v>[^<]*<\/c:v>\s*<\/c:pt>/g,
    "",
  );

  /*
   * Atualiza ptCount.
   */
  if (
    /<c:ptCount\s+val="\d+"\s*\/>/.test(
      middle,
    )
  ) {
    middle = middle.replace(
      /<c:ptCount\s+val="\d+"\s*\/>/,
      '<c:ptCount val="' +
        values.length +
        '"/>',
    );
  }

  /*
   * Insere os novos pontos.
   */
  const ptCountIdx =
    middle.indexOf(
      "<c:ptCount",
    );

  if (ptCountIdx >= 0) {
    const insertAt =
      middle.indexOf(
        "/>",
        ptCountIdx,
      ) + 2;

    middle =
      middle.slice(
        0,
        insertAt,
      ) +
      pts +
      middle.slice(
        insertAt,
      );
  } else {
    middle += pts;
  }

  return (
    xml.slice(0, start) +
    middle +
    xml.slice(end)
  );
}

function dateToSerial(date: Date): number {
  const excelEpoch = Date.UTC(1899, 11, 30);

  const utcDate = Date.UTC(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
  );

  return Math.floor(
    (utcDate - excelEpoch) / 86400000,
  );
}


/* =========================================================
   GERADOR PRINCIPAL
========================================================= */

export async function gerarApresentacao(
  files: GeneratorFiles,
  onLog: (msg: string, cls?: LogKind) => void = () => {},
): Promise<Blob> {
  log = onLog;

  log("Iniciando geração...");

  /*
   * ============================================================
   * LEITURA DOS ARQUIVOS
   * ============================================================
   *
   * Cada arquivo agora é opcional.
   *
   * Se o arquivo existir:
   *   -> lê e calcula os dados.
   *
   * Se não existir:
   *   -> cria dados vazios/default.
   *
   * Dessa forma é possível gerar a apresentação mesmo
   * enviando somente 1 arquivo.
   */

  let wbSm: any = null;
  let wbAs: any = null;
  let wbAc: any = null;
  let wbFv: any = null;
  let wbFe: any = null;
  let wbSa: any = null;

  /*
   * Score Mensal
   */
  if (files.sm) {
    log("Lendo Score Mensal...");
    wbSm = await readXlsx(files.sm, "Score Mensal");
  } else {
    log("Score Mensal não enviado. Usando valores padrão.", "warn");
  }

  /*
   * Associações
   */
  if (files.as) {
    log("Lendo Associações...");
    wbAs = await readXlsx(files.as, "Associações");
  } else {
    log("Associações não enviada. Usando valores padrão.", "warn");
  }

  /*
   * Validação
   */
  if (files.ac) {
    log("Lendo Validação...");
    wbAc = await readXlsx(files.ac, "Validação");
  } else {
    log("Validação não enviada. Usando valores padrão.", "warn");
  }

  /*
   * Frequência Vendas
   */
  if (files.fv) {
    log("Lendo Frequência Vendas...");
    wbFv = await readXlsx(files.fv, "Frequência Vendas");
  } else {
    log("Frequência Vendas não enviada. Usando valores padrão.", "warn");
  }

  /*
   * Frequência Estoque
   */
  if (files.fe) {
    log("Lendo Frequência Estoque...");
    wbFe = await readXlsx(files.fe, "Frequência Estoque");
  } else {
    log("Frequência Estoque não enviada. Usando valores padrão.", "warn");
  }

  /*
   * Score Fechado
   */
  if (files.sa) {
    log("Lendo Score Fechado...");
    wbSa = await readXlsx(files.sa, "Score Fechado");
  } else {
    log("Score Fechado não enviado. Usando valores padrão.", "warn");
  }

  /*
   * ============================================================
   * PARSING
   * ============================================================
   */

  const sm = wbSm
    ? parseScoreMensal(wbSm)
    : [];

  const as = wbAs
    ? parseAssoc(wbAs)
    : {
        distribuidores: 0,
        industrias: 0,
        associados: 0,
        naoIdent: 0,
        rejeitados: 0,
        impactados: 0,
      };

  const val = wbAc
    ? parseValidacao(wbAc)
    : {
        latest: null,
        total: 0,
        riscadas: 0,
        dentro: 0,
        fora: 0,
        pend: 0,
        div: 0,
        tabela: [],
        pctDentro: 0,
        pctFora: 0,
        pctPend: 0,
        pctDiv: 0,
        pctValidado: 0,
      };

  const fv = wbFv
    ? parseFrequenciaDias(wbFv)
    : {
        mesLabel: null,
        list: [],
        atualizados: 0,
        atencao: 0,
        ofensores: 0,
        total: 0,
        ofensoresList: [],
        pctFreq: 0,
      };

  const fe = wbFe
    ? parseFrequenciaDias(wbFe)
    : {
        mesLabel: null,
        list: [],
        atualizados: 0,
        atencao: 0,
        ofensores: 0,
        total: 0,
        ofensoresList: [],
        pctFreq: 0,
      };

  const sf = wbSa
    ? parseScoreFechado(wbSa)
    : {
        industria: null,
        score: 0,
        frequencia: 0,
        acuracidade: 0,
        consistencia: 0,
        validacao: 0,
      };

  /*
   * ============================================================
   * LINHAS RISCADAS
   * ============================================================
   */

  const worksheets = [
    wbSm,
    wbAs,
    wbAc,
    wbFv,
    wbFe,
    wbSa,
  ].filter(Boolean);

  const riscTotal = worksheets.reduce(
    (total, wb) => total + totalStruck(wb),
    0,
  );



 
  /*
   * ============================================================
   * DATAS
   * ============================================================
   */

  const hoje = new Date();

const mesAtualCap = MESES_PT[hoje.getMonth()];

const anoAtual = String(hoje.getFullYear());

  const dataAnterior = new Date(
    hoje.getFullYear(),
    hoje.getMonth() - 1,
    1,
  );

  const mesCap = cap(
    MESES_PT[dataAnterior.getMonth()],
  );

  const anoRef = String(
    dataAnterior.getFullYear(),
  );

  const mesNome =
    mesCap + " " + anoRef;

  log(
    "Período de referência: " + mesNome,
    "ok",
  );

  /*
   * ============================================================
   * TEMPLATE
   * ============================================================
   */

  log("Carregando template...");

  const zip = await JSZip.loadAsync(
    await loadTemplateBuffer(),
  );

  /*
   * ============================================================
   * SLIDE 3 — CONSOLIDADO
   * ============================================================
   */

  log("Editando Consolidado...");

  let s3 = await zip
    .file("ppt/slides/slide3.xml")
    ?.async("string");

  if (s3) {
    s3 = replaceRun(
      s3,
      "0%",
      " " + fmtPctBR(sf.score),
    );

    s3 = replaceRun(
      s3,
      " 0%",
      " " + fmtPctBR(sf.frequencia),
    );

    s3 = replaceRun(
      s3,
      " 0%",
      " " + fmtPctBR(sf.validacao),
    );

    s3 = replaceRun(
      s3,
      " 0%",
      " " + fmtPctBR(sf.acuracidade),
    );

    s3 = replaceRun(
      s3,
      " 0%",
      " " + fmtPctBR(sf.consistencia),
    );

    zip.file(
      "ppt/slides/slide3.xml",
      s3,
    );
  }

  /*
   * ============================================================
   * GRÁFICO
   * ============================================================
   *
   * Só atualiza se houver dados de Score Mensal.
   */

  if (sm.length > 0) {
    log(
      "Atualizando gráfico de evolução...",
    );

    const chartFile =
      zip.file("ppt/charts/chart1.xml");

    if (chartFile) {
      let c1 =
        await chartFile.async("string");

      c1 = setPointsInSection(
        c1,
        "<c:cat>",
        "</c:cat>",
        sm.map((r) =>
          dateToSerial(r.date),
        ),
      );

      c1 = setPointsInSection(
        c1,
        "<c:val>",
        "</c:val>",
        sm.map((r) => r.score),
      );

      zip.file(
        "ppt/charts/chart1.xml",
        c1,
      );

      log(
        "Gráfico de evolução: " +
          sm.length +
          " meses",
        "ok",
      );
    }
  } else {
    log(
      "Score Mensal não enviado. Gráfico mantido sem alteração.",
      "warn",
    );
  }

  /*
   * ============================================================
   * SLIDE 4 — VALIDAÇÃO MENSAL
   * ============================================================
   */

  log(
    "Editando Validação Mensal...",
  );

  let s4 = await zip
    .file("ppt/slides/slide4.xml")
    ?.async("string");

  if (s4) {
    s4 = fillMesAno(
      s4,
      mesCap,
      anoRef,
    );

    s4 = replaceNumberAfterLabel(
  s4,
  "Dentro do Prazo",
  pctInt(val.pctDentro),
);

s4 = replaceNumberAfterLabel(
  s4,
  "Pendentes",
  pctInt(val.pctPend),
);

s4 = replaceNumberAfterLabel(
  s4,
  "Divergentes",
  pctInt(val.pctDiv),
);

s4 = replaceNumberAfterLabel(
  s4,
  "Fora do prazo",
  pctInt(val.pctFora),
);

    zip.file(
      "ppt/slides/slide4.xml",
      s4,
    );

    await fillValidacaoSlide(
      zip,
      4,
      val.tabela.map((r) => [
        r.cnpj,
        r.nome,
        r.status,
      ]),
    );
  }

  /*
   * ============================================================
   * FREQUÊNCIA
   * ============================================================
   */

  log(
    "Editando slides de Frequência...",
  );

  async function fillFreq(
    slideNum: number,
    f: any,
    nome: string,
  ) {
    const file = zip.file(
      "ppt/slides/slide" +
        slideNum +
        ".xml",
    );

    if (!file) return;

    let s =
      await file.async("string");

    s = fillMesAno(
      s,
      mesAtualCap,
      anoAtual,
    );

    const t = f.total || 1;

    s = replaceRunsSeq(
      s,
      "%",
      [
        fmtPctBR(
          f.atualizados / t,
        ),
        fmtPctBR(
          f.atencao / t,
        ),
        fmtPctBR(
          f.ofensores / t,
        ),
      ],
    );

    s = replaceRunsSeq(
      s,
      "777",
      [
        String(f.atualizados),
        String(f.atencao),
        String(f.ofensores),
      ],
    );

    zip.file(
      "ppt/slides/slide" +
        slideNum +
        ".xml",
      s,
    );

  
  }

  await fillFreq(
    6,
    fv,
    "Frequência de Vendas",
  );

  await fillFreq(
    8,
    fe,
    "Frequência de Estoque",
  );

  /*
   * ============================================================
   * OFENSORES
   * ============================================================
   */

  log(
    "Editando slides de Ofensores...",
  );

  const pctFreqOf = (
    d: any,
  ) => {
    const totalDias =
      Number(d.considerados) || 0;

    const diasSemDados =
      Number(d.ausentes) || 0;

    return totalDias > 0
      ? diasSemDados / totalDias
      : 0;
  };

  const mapOf = (
    d: any,
  ) => [
    d.cnpj,
    d.nome,
    String(d.ausentes),
    fmtNumBR(
      pctFreqOf(d) * 100,
    ) + "%",
  ];

  const dataMes = (
    xml: string,
  ) =>
    fillMesAno(
      xml,
      mesAtualCap,
      anoAtual,
    );

  await fillTableAcrossSlides(
    zip,
    [7],
    fv.ofensoresList.map(mapOf),
    "Ofensores Sellout",
    dataMes,
  );

  await fillTableAcrossSlides(
    zip,
    [9],
    fe.ofensoresList.map(mapOf),
    "Ofensores Estoque",
    dataMes,
  );

  /*
   * ============================================================
   * SLIDE 10 — INCONSISTÊNCIAS
   * ============================================================
   */

  log(
    "Editando Tratamento de Inconsistências...",
  );

  let s10 = await zip
    .file("ppt/slides/slide10.xml")
    ?.async("string");

  if (s10) {
    s10 = replaceNumberBeforeLabel(
      s10,
      "Distribuidores  Identificados",
      as.distribuidores,
    );

    s10 = replaceNumberBeforeLabel(
      s10,
      "Produtos Impactados",
      as.impactados,
    );

    s10 = replaceNumberBeforeLabel(
      s10,
      "Produtos Associados",
      as.associados,
    );

    s10 = replaceNumberBeforeLabel(
      s10,
      "Produtos Não Identificados",
      as.naoIdent,
    );

    s10 = replaceNumberBeforeLabel(
      s10,
      "Produtos Rejeitados",
      as.rejeitados,
    );

    zip.file(
      "ppt/slides/slide10.xml",
      s10,
    );
  }

  /*
   * ============================================================
   * GERAÇÃO DO PPTX
   * ============================================================
   */

  log(
    "Compactando arquivo...",
  );

  const out =
    await zip.generateAsync({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      compression: "DEFLATE",
    });

  log(
    "✓ Apresentação gerada.",
    "ok",
  );

  return out as Blob;
}