import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useRef, useState } from "react";
import {
  gerarApresentacao,
  type GeneratorFiles,
  type LogKind,
} from "@/lib/pptx-generator";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Gerador de Indicadores Data Quality em PPTX" },
      {
        name: "description",
        content:
          "Envie as planilhas de frequência, validação, associações e score e gere a apresentação de Data Quality no template oficial.",
      },
      { property: "og:title", content: "Gerador de Indicadores Data Quality" },
      {
        property: "og:description",
        content:
          "Transforme planilhas de qualidade de dados em uma apresentação PPTX padronizada, pronta para apresentar.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

type Slot = {
  id: keyof GeneratorFiles;
  titulo: string;
  fonte: string;
  ajudaTitulo: string;
  ajudaTexto: string;
  link: string;
  linkTexto: string;
};

const SLOTS: Slot[] = [
  {
    id: "fv",
    titulo: "Frequência de Vendas",
    fonte: "via Toolbelt",
    ajudaTitulo: "Frequência de Vendas",
    ajudaTexto:
      "Arquivo extraído pelo Toolbelt. Filtre o mês e a conexão antes de realizar a extração.",
    link: "https://mgmt.neogrid.com/toolbelt/indirect/di/data-panorama/?movement=1&month=2026-07-03&vision=585e9fb1-ae9c-44ed-c7a0-08d942214fee",
    linkTexto: "Link para extração →",
  },
  {
    id: "fe",
    titulo: "Frequência de Estoque",
    fonte: "via Toolbelt",
    ajudaTitulo: "Frequência de Estoque",
    ajudaTexto:
      "Arquivo extraído pelo Toolbelt. Filtre o mês e a conexão antes de realizar a extração.",
    link: "https://mgmt.neogrid.com/toolbelt/indirect/di/data-panorama/?movement=2&month=2026-07-03&vision=585e9fb1-ae9c-44ed-c7a0-08d942214fee",
    linkTexto: "Link de extração →",
  },
  {
    id: "ac",
    titulo: "Validação mensal",
    fonte: "via Toolbelt",
    ajudaTitulo: "Validação mensal",
    ajudaTexto: "Arquivo utilizado para a validação mensal dos dados.",
    link: "https://mgmt.neogrid.com/toolbelt/indirect/di/accuracy-indicator/?month=2026-07-03&vision=585e9fb1-ae9c-44ed-c7a0-08d942214fee&status=1",
    linkTexto: "Link de extração →",
  },
  {
    id: "as",
    titulo: "Associações",
    fonte: "Planilha de associação",
    ajudaTitulo: "Associações",
    ajudaTexto:
      "Planilha de associação obtida com o analista responsável pelo SharePoint da célula.",
    link: "#",
    linkTexto: "Falar com analista responsável ou SharePoint",
  },
  {
    id: "sm",
    titulo: "Score mensal (evolução)",
    fonte: "via Power BI",
    ajudaTitulo: "Score mensal (evolução)",
    ajudaTexto: "Dados utilizados para acompanhar a evolução do score mensal.",
    link: "https://app.powerbi.com/groups/ba163989-3229-4868-ba80-d85b709d568a/reports/7789f034-b523-416a-84bc-f91f2324eb52/d92366b721abca391534?experience=power-bi",
    linkTexto: "Link de extração →",
  },
];

function fmtSize(bytes: number) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

function UploadItem({
  slot,
  file,
  onPick,
}: {
  slot: Slot;
  file: File | null;
  onPick: (f: File | null) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  const [help, setHelp] = useState(false);

  const aceita = (f?: File | null) => {
    if (f && /\.xlsx$/i.test(f.name)) onPick(f);
  };

  return (
    <div
      className={
        "upload-item" + (file ? " has-file" : "") + (drag ? " drag-over" : "")
      }
      onDragOver={(e) => {
        e.preventDefault();
        setDrag(true);
      }}
      onDragLeave={() => setDrag(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDrag(false);
        aceita(e.dataTransfer.files?.[0]);
      }}
    >
      <div className="upload-icon">
        <svg
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
          <polyline points="17 8 12 3 7 8" />
          <line x1="12" y1="3" x2="12" y2="15" />
        </svg>
      </div>

      <div className="upload-info">
        <div className="upload-title-row">
          <div className="upload-title">
            <div>
              <strong>{slot.titulo}</strong>
              <small>{slot.fonte}</small>
            </div>
          </div>

          <button
            type="button"
            className="help-button"
            aria-expanded={help}
            aria-label={"Ajuda sobre " + slot.titulo}
            onClick={() => setHelp((v) => !v)}
          >
            ?
          </button>

          <div
            className={"upload-help help-source-popover" + (help ? " is-open active" : "")}
          >
            <div className="help-source-header">
              <span>?</span>
              <strong>Origem do arquivo</strong>
            </div>
            <strong>{slot.ajudaTitulo}</strong>
            <p>{slot.ajudaTexto}</p>
            <a
              href={slot.link}
              target="_blank"
              rel="noopener"
              className="help-source-link"
            >
              {slot.linkTexto}
            </a>
          </div>
        </div>

        {file ? (
          <>
            <div className="upload-file-name" title={file.name}>
              {file.name}
            </div>
            <div className="file-size">{fmtSize(file.size)}</div>
            <div className="file-status">Arquivo anexado</div>
            <div className="file-actions">
              <button
                type="button"
                className="file-action replace"
                onClick={() => inputRef.current?.click()}
              >
                Trocar arquivo
              </button>
              <button
                type="button"
                className="file-action remove"
                onClick={() => {
                  onPick(null);
                  if (inputRef.current) inputRef.current.value = "";
                }}
              >
                Remover
              </button>
            </div>
          </>
        ) : (
          <div className="upload-area" onClick={() => inputRef.current?.click()}>
            <div className="upload-state empty">
              <strong>Nenhum arquivo selecionado</strong>
              <span>
                Clique aqui para selecionar ou arraste o arquivo .xlsx até este
                card
              </span>
            </div>
          </div>
        )}

        <input
          ref={inputRef}
          type="file"
          accept=".xlsx"
          style={{ display: "none" }}
          onChange={(e) => aceita(e.target.files?.[0] ?? null)}
        />
      </div>
    </div>
  );
}

function Index() {
  const [files, setFiles] = useState<Record<string, File | null>>({});
  const [logs, setLogs] = useState<{ msg: string; cls: LogKind }[]>([]);
  const [busy, setBusy] = useState(false);

  const enviados = SLOTS.filter((s) => files[s.id]).length;
  const pct = Math.round((enviados / SLOTS.length) * 100);
  const pronto = useMemo(() => SLOTS.every((s) => files[s.id]), [files]);

  async function handleGerar() {
    setBusy(true);
    setLogs([]);
    const push = (msg: string, cls?: LogKind) =>
      setLogs((prev) => [...prev, { msg, cls }]);
    try {
      const blob = await gerarApresentacao(
        files as unknown as GeneratorFiles,
        push,
      );
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "DataQuality_gerado.pptx";
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      push("Download iniciado.", "ok");
    } catch (e) {
      push("Erro: " + ((e as Error)?.message ?? String(e)), "err");
      console.error(e);
    } finally {
      setBusy(false);
    }
  }

  const handleBaixarModelo = () => {
  const link = document.createElement("a");
  link.href = "/templateFinal.pptx";
  link.download = "templateFinal.pptx";
  document.body.appendChild(link);
  link.click();
  link.remove();
};
  return (
    <>
      <div className="topbar">
        <span className="brand">
          Kairos
          <span
            style={{ opacity: 0.6, textTransform: "none", fontWeight: 500 }}
          >
            {" "}
            · Data Quality
          </span>
        </span>
      </div>

      <div className="wrap">
        <h1>Indicadores — Data Quality</h1>
        <div className="sub">
          Transforme seus arquivos de dados em apresentações PPTX semanais,
          quinzenais ou mensais. O conteúdo é organizado automaticamente em um
          template padronizado, pronto para apresentação.
        </div>

        <div className="card">
          <div className="upload-progress-container">
            <div className="progress-info">
              <span>
                {enviados} de {SLOTS.length} arquivos enviados
              </span>
              <span id="progress-percent">{pct}%</span>
            </div>
            <div className="progress-bar">
              <div id="progress-fill" style={{ width: pct + "%" }} />
            </div>
          </div>

          <div className="row">
            {SLOTS.map((slot) => (
              <UploadItem
                key={slot.id}
                slot={slot}
                file={files[slot.id] ?? null}
                onPick={(f) => setFiles((prev) => ({ ...prev, [slot.id]: f }))}
              />
            ))}
          </div>


<div className="flex gap-4">
          <button
            className="btn"
            type="button"
            disabled={!pronto || busy}
            onClick={handleGerar}
          >
            {busy ? "Gerando apresentação..." : "Gerar PPT"}
          </button>

        

<button
  className="btn"
  type="button"
  onClick={handleBaixarModelo}
>
  Baixar modelo PPTX
</button>
          </div>

          {logs.length > 0 && (
            <div id="log">
              {logs.map((l, i) => (
                <div key={i} className={l.cls}>
                  {l.msg}
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="cards-row">
          <div className="card support-card">
            <h2>Material de apoio</h2>
            <p className="hint" style={{ marginTop: 0 }}>
              Documento com as orientações de uso e preenchimento do material.
            </p>
            <a
              className="apoio"
              href="https://neogridsoftware-my.sharepoint.com/:w:/r/personal/branca_azevedo_neogrid_com/_layouts/15/doc2.aspx?sourcedoc=%7B3A323CB5-F30D-449D-BABD-5287E7A2C90A%7D&file=Document%206.docx&action=editNew&mobileredirect=true"
              target="_blank"
              rel="noopener"
            >
              Abrir material de apoio
            </a>
          </div>
        </div>
      </div>

      <footer>
        Feito e idealizado por Branca Azevedo, quaisquer dúvidas entrar em
        contato.
      </footer>
    </>
  );
}
