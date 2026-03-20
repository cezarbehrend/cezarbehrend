/* Dashboard local de veículos de mídia (Rádios/TVs) - sem backend */
(() => {
  const COOPS = ["3001", "3003", "3007", "3008", "3010", "3260"];
  const STATE = {
    rawRadio: [],
    rawTv: [],
    vehicles: [],
    filtered: [],
    mapDataByUf: {},
    tableRows: [],
    sort: { key: "municipio", asc: true }
  };

  const el = (id) => document.getElementById(id);
  const dom = {
    excelInput: el("excelInput"),
    reloadBtn: el("reloadBtn"),
    clearFiltersBtn: el("clearFiltersBtn"),
    statusBar: el("statusBar"),
    filterUf: el("filterUf"),
    filterTipo: el("filterTipo"),
    filterCoop: el("filterCoop"),
    filterVeiculo: el("filterVeiculo"),
    filterMunicipio: el("filterMunicipio"),
    kpiGrid: el("kpiGrid"),
    analyticsContent: el("analyticsContent"),
    tableHead: document.querySelector("#detailsTable thead"),
    tableBody: document.querySelector("#detailsTable tbody"),
    tableSearch: el("tableSearch"),
    exportCsvBtn: el("exportCsvBtn")
  };

  let map;
  let geoLayer;

  init();

  function init() {
    initMap();
    setStatus("Aguardando planilha .xlsx", "");
    fillCoopSelect();
    bindEvents();
  }

  function bindEvents() {
    dom.excelInput.addEventListener("change", onExcelUpload);
    dom.reloadBtn.addEventListener("click", () => dom.excelInput.click());
    dom.clearFiltersBtn.addEventListener("click", clearFilters);
    [dom.filterUf, dom.filterTipo, dom.filterCoop, dom.filterVeiculo].forEach((node) =>
      node.addEventListener("change", applyFilters)
    );
    dom.filterMunicipio.addEventListener("input", applyFilters);
    dom.tableSearch.addEventListener("input", renderTable);
    dom.exportCsvBtn.addEventListener("click", exportCsv);
  }

  function initMap() {
    map = L.map("map", { zoomControl: true }).setView([-22.2, -42.8], 6);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      attribution: "&copy; OpenStreetMap contributors"
    }).addTo(map);
  }

  async function onExcelUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setStatus("Processando planilha...", "");
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      validateWorkbook(workbook);

      STATE.rawRadio = parseRaiosSheet(workbook.Sheets["Rádios"]);
      STATE.rawTv = parseTvsSheet(workbook.Sheets["TVs"]);
      STATE.vehicles = [...STATE.rawRadio, ...STATE.rawTv];

      await loadGeoData();
      fillUfSelect();
      fillVehicleSelect(STATE.vehicles);
      applyFilters();
      setStatus("Planilha carregada com sucesso.", "success");
    } catch (err) {
      console.error(err);
      setStatus(`Erro ao carregar planilha: ${err.message}`, "error");
    }
  }

  function validateWorkbook(wb) {
    if (!wb.Sheets["Rádios"] || !wb.Sheets["TVs"]) {
      throw new Error("A planilha deve conter as abas 'Rádios' e 'TVs'.");
    }
  }

  function parseRaiosSheet(sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    return rows.map((row, i) => {
      const n = normalizeKeys(row);
      const uf = cleanText(n.uf || "");
      const sedeMunicipio = cleanCity(n.municipio_sede || "");
      const veiculo = cleanText(n.nome_da_emissora || `Rádio ${i + 1}`);
      const cobertura = parseCoverageCities(n.cidades_de_cobertura_sinal_principal || "", uf);
      const coopAtivasLinha = activeCoopsFromRow(n);

      const sedeEntry = {
        municipio: sedeMunicipio,
        uf,
        papel: "sede",
        populacao: toNumber(n.populacao_sede_hab),
        associados: toNumber(n.qtde_associados_no_municipio_sede),
        cooperativas: coopAtivasLinha
      };

      const coverageEntries = cobertura.map((c) => ({
        municipio: c.municipio,
        uf: c.uf || uf,
        papel: normalize(c.municipio) === normalize(sedeMunicipio) ? "sede" : "cobertura",
        populacao: null,
        associados: null,
        cooperativas: coopAtivasLinha
      }));

      return {
        id: `radio-${i}`,
        tipo: "radio",
        uf,
        nome: veiculo,
        sedeMunicipio: sedeMunicipio || "Não informado",
        sedePopulacao: toNumber(n.populacao_sede_hab),
        sedeAssociados: toNumber(n.qtde_associados_no_municipio_sede),
        totalPopulacao: toNumber(n.populacao_total_de_todas_cidades_com_sinal),
        municipios: dedupeMunicipios([sedeEntry, ...coverageEntries]),
        cooperativasArea: coopAtivasLinha
      };
    });
  }

  function parseTvsSheet(sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(normalizeKeys);
    const byExibidora = new Map();
    rows.forEach((row) => {
      const exibidora = cleanText(row.exibidora || "TV não identificada");
      if (!byExibidora.has(exibidora)) byExibidora.set(exibidora, []);
      byExibidora.get(exibidora).push(row);
    });

    let idx = 0;
    return Array.from(byExibidora.entries()).map(([exibidora, group]) => {
      const municipios = dedupeMunicipios(group.map((r) => ({
        municipio: cleanCity(r.municipio || ""),
        uf: cleanText(r.uf || ""),
        papel: "cobertura_tv",
        populacao: toNumber(r.populacao),
        associados: null,
        cooperativas: activeCoopsFromRow(r)
      })));

      const coopArea = [...new Set(municipios.flatMap((m) => m.cooperativas))];
      const ufSet = [...new Set(municipios.map((m) => m.uf).filter(Boolean))];
      return {
        id: `tv-${idx++}`,
        tipo: "tv",
        uf: ufSet.length === 1 ? ufSet[0] : (ufSet[0] || ""),
        nome: exibidora,
        sedeMunicipio: "Não informado",
        sedePopulacao: null,
        sedeAssociados: null,
        totalPopulacao: municipios.reduce((s, m) => s + (m.populacao || 0), 0),
        municipios,
        cooperativasArea: coopArea
      };
    });
  }

  async function loadGeoData() {
    if (STATE.mapDataByUf.ES && STATE.mapDataByUf.RJ) return;
    // GeoJSON simplificado para municípios do ES e RJ
    const [es, rj] = await Promise.all([
      fetch("https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-32-mun.json").then((r) => r.json()),
      fetch("https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-33-mun.json").then((r) => r.json())
    ]);
    STATE.mapDataByUf.ES = es;
    STATE.mapDataByUf.RJ = rj;
  }

  function fillCoopSelect() {
    COOPS.forEach((c) => {
      dom.filterCoop.insertAdjacentHTML("beforeend", `<option value="${c}">${c}</option>`);
    });
  }

  function fillUfSelect() {
    const ufs = [...new Set(STATE.vehicles.map((v) => v.uf).filter(Boolean))].sort();
    dom.filterUf.innerHTML = '<option value="">Todos</option>' + ufs.map((u) => `<option value="${u}">${u}</option>`).join("");
  }

  function fillVehicleSelect(list) {
    const names = [...new Set(list.map((v) => v.nome))].sort((a, b) => a.localeCompare(b));
    dom.filterVeiculo.innerHTML = '<option value="">Todos</option>' + names.map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join("");
  }

  function applyFilters() {
    const f = {
      uf: dom.filterUf.value,
      tipo: dom.filterTipo.value,
      coop: dom.filterCoop.value,
      veiculo: dom.filterVeiculo.value,
      municipio: normalize(dom.filterMunicipio.value)
    };

    STATE.filtered = STATE.vehicles.filter((v) => {
      if (f.uf && v.uf !== f.uf) return false;
      if (f.tipo && v.tipo !== f.tipo) return false;
      if (f.veiculo && v.nome !== f.veiculo) return false;
      if (f.coop && !v.cooperativasArea.includes(f.coop) && !v.municipios.some((m) => m.cooperativas.includes(f.coop))) return false;
      if (f.municipio && !v.municipios.some((m) => normalize(m.municipio).includes(f.municipio))) return false;
      return true;
    });

    fillVehicleSelect(STATE.filtered.length ? STATE.filtered : STATE.vehicles);
    const selectedVehicle = selectVehiclePriority(STATE.filtered);
    const detailsRows = selectedVehicle ? selectedVehicle.municipios : flattenMunicipios(STATE.filtered);

    STATE.tableRows = detailsRows.map((m) => ({
      municipio: m.municipio,
      papel: m.papel,
      populacao: m.populacao,
      cooperativas: m.cooperativas.join(", "),
      associados: m.associados,
      veiculo: selectedVehicle?.nome || findVehicleForMunicipio(STATE.filtered, m),
      uf: m.uf
    }));

    renderKpis(selectedVehicle, STATE.filtered);
    renderMap(selectedVehicle, STATE.filtered);
    renderTable();
    renderAnalytics(selectedVehicle, STATE.filtered);
  }

  function selectVehiclePriority(list) {
    const explicit = dom.filterVeiculo.value;
    if (explicit) return list.find((v) => v.nome === explicit) || null;
    return list.length === 1 ? list[0] : null;
  }

  function renderKpis(vehicle, set) {
    if (!set.length) {
      dom.kpiGrid.innerHTML = `<div class="card empty">Sem dados para os filtros selecionados.</div>`;
      return;
    }
    const scope = vehicle ? [vehicle] : set;
    const allMun = flattenMunicipios(scope);
    const totalPop = vehicle ? vehicle.totalPopulacao : scope.reduce((s, v) => s + (v.totalPopulacao || 0), 0);
    const coop = [...new Set(allMun.flatMap((m) => m.cooperativas))];
    const sede = vehicle ? vehicle.sedeMunicipio : "Selecione um veículo";

    const cards = [
      ["Nome do veículo", vehicle?.nome || `${set.length} veículos`],
      ["Tipo de veículo", vehicle ? (vehicle.tipo === "radio" ? "Rádio" : "TV") : "Múltiplos"],
      ["Estado", vehicle?.uf || [...new Set(set.map((v) => v.uf))].join(", ") || "-"],
      ["Município sede", sede || "Não informado"],
      ["População da sede", fmt(vehicle?.sedePopulacao)],
      ["População impactada", fmt(totalPop)],
      ["Associados na sede", fmt(vehicle?.sedeAssociados)],
      ["Municípios cobertos", fmt(new Set(allMun.map((m) => normalize(m.municipio) + m.uf)).size)],
      ["Cooperativas na cobertura", coop.join(", ") || "-"],
      ["Qtd. de cooperativas", fmt(coop.length)]
    ];

    dom.kpiGrid.innerHTML = cards
      .map(([title, value]) => `<article class="kpi"><h3>${escapeHtml(title)}</h3><p>${escapeHtml(String(value ?? "-"))}</p></article>`)
      .join("");
  }

  function renderMap(vehicle, set) {
    if (geoLayer) geoLayer.remove();
    if (!set.length) return;

    const uf = vehicle?.uf || dom.filterUf.value || set[0].uf;
    const geo = STATE.mapDataByUf[uf];
    if (!geo) return;

    const munMap = new Map();
    const source = vehicle ? vehicle.municipios : flattenMunicipios(set);
    source.forEach((m) => munMap.set(normalize(m.municipio), m));

    geoLayer = L.geoJSON(geo, {
      style: (feature) => {
        const name = getFeatureMunicipio(feature);
        const m = munMap.get(normalize(name));
        let fill = "#dee6f3";
        if (m) fill = m.papel === "sede" ? "#f08a24" : "#52a7ff";
        return { color: "#9cb0cb", weight: 1, fillColor: fill, fillOpacity: m ? 0.85 : 0.45 };
      },
      onEachFeature: (feature, layer) => {
        const name = getFeatureMunicipio(feature);
        const m = munMap.get(normalize(name));
        const papel = m?.papel || "fora da seleção";
        const cooperativas = m?.cooperativas?.length ? m.cooperativas.join(", ") : "-";
        layer.bindTooltip(`
          <strong>${escapeHtml(name)}</strong><br>
          Papel: ${escapeHtml(papel)}<br>
          População: ${escapeHtml(fmt(m?.populacao))}<br>
          Cooperativas: ${escapeHtml(cooperativas)}<br>
          Associados sede: ${escapeHtml(fmt(m?.associados))}
        `);
      }
    }).addTo(map);

    const bounds = geoLayer.getBounds();
    if (bounds.isValid()) map.fitBounds(bounds.pad(0.05));
  }

  function renderTable() {
    const headers = ["municipio", "papel", "populacao", "cooperativas", "associados", "veiculo", "uf"];
    dom.tableHead.innerHTML = `<tr>${headers.map((h) => `<th data-key="${h}">${h.toUpperCase()}</th>`).join("")}</tr>`;

    dom.tableHead.querySelectorAll("th").forEach((th) => {
      th.addEventListener("click", () => {
        const key = th.dataset.key;
        if (STATE.sort.key === key) STATE.sort.asc = !STATE.sort.asc;
        else STATE.sort = { key, asc: true };
        renderTable();
      });
    });

    const q = normalize(dom.tableSearch.value || "");
    const filtered = STATE.tableRows.filter((r) => !q || Object.values(r).some((v) => normalize(String(v || "")).includes(q)));
    const sorted = [...filtered].sort((a, b) => sortCompare(a[STATE.sort.key], b[STATE.sort.key], STATE.sort.asc));

    dom.tableBody.innerHTML = sorted
      .map((r) => `<tr><td>${escapeHtml(r.municipio || "-")}</td><td>${escapeHtml(r.papel || "-")}</td><td>${escapeHtml(fmt(r.populacao))}</td><td>${escapeHtml(r.cooperativas || "-")}</td><td>${escapeHtml(fmt(r.associados))}</td><td>${escapeHtml(r.veiculo || "-")}</td><td>${escapeHtml(r.uf || "-")}</td></tr>`)
      .join("") || `<tr><td colspan="7" class="empty">Sem linhas para exibir.</td></tr>`;
  }

  function renderAnalytics(vehicle, set) {
    if (!set.length) {
      dom.analyticsContent.innerHTML = `<p class="empty">Sem dados analíticos para os filtros atuais.</p>`;
      return;
    }

    const scope = vehicle ? [vehicle] : set;
    const all = flattenMunicipios(scope);
    const totalPop = scope.reduce((s, v) => s + (v.totalPopulacao || 0), 0);

    const coopAgg = {};
    all.forEach((m) => {
      const pop = m.populacao || 0;
      m.cooperativas.forEach((c) => {
        if (!coopAgg[c]) coopAgg[c] = { municipios: new Set(), pop: 0 };
        coopAgg[c].municipios.add(`${normalize(m.municipio)}|${m.uf}`);
        coopAgg[c].pop += pop;
      });
    });

    const coopRank = Object.entries(coopAgg)
      .map(([coop, obj]) => ({ coop, municipios: obj.municipios.size, pop: obj.pop, pct: totalPop ? (obj.pop / totalPop) * 100 : 0 }))
      .sort((a, b) => b.pop - a.pop);

    const cityRank = [...all]
      .sort((a, b) => (b.populacao || 0) - (a.populacao || 0))
      .slice(0, 10);

    dom.analyticsContent.innerHTML = `
      <article class="analytics-box">
        <h3>Resumo geral</h3>
        <ul class="list">
          <li>Municípios cobertos: <strong>${new Set(all.map((m) => normalize(m.municipio) + m.uf)).size}</strong></li>
          <li>População impactada: <strong>${fmt(totalPop)}</strong></li>
          <li>Sede do veículo: <strong>${vehicle?.sedeMunicipio || "Selecione um veículo"}</strong></li>
          <li>Associados da sede: <strong>${fmt(vehicle?.sedeAssociados)}</strong></li>
          <li>Cooperativas na área: <strong>${[...new Set(all.flatMap((m) => m.cooperativas))].join(", ") || "-"}</strong></li>
        </ul>
      </article>
      <article class="analytics-box">
        <h3>Ranking cooperativas por população coberta</h3>
        <ol class="list">${coopRank.map((c) => `<li>${c.coop}: ${fmt(c.pop)} hab. • ${c.municipios} mun. • ${c.pct.toFixed(1)}%</li>`).join("") || "<li>-</li>"}</ol>
      </article>
      <article class="analytics-box">
        <h3>Ranking cidades por população</h3>
        <ol class="list">${cityRank.map((c) => `<li>${escapeHtml(c.municipio)} (${c.uf}): ${fmt(c.populacao)}</li>`).join("") || "<li>-</li>"}</ol>
      </article>
    `;
  }

  function flattenMunicipios(vehicles) {
    return vehicles.flatMap((v) => v.municipios.map((m) => ({ ...m, veiculo: v.nome })));
  }

  function findVehicleForMunicipio(vehicles, m) {
    const hit = vehicles.find((v) => v.municipios.some((x) => normalize(x.municipio) === normalize(m.municipio) && x.uf === m.uf));
    return hit?.nome || "-";
  }

  function clearFilters() {
    dom.filterUf.value = "";
    dom.filterTipo.value = "";
    dom.filterCoop.value = "";
    dom.filterVeiculo.value = "";
    dom.filterMunicipio.value = "";
    dom.tableSearch.value = "";
    applyFilters();
  }

  function exportCsv() {
    if (!STATE.tableRows.length) return;
    const headers = ["municipio", "papel", "populacao", "cooperativas", "associados", "veiculo", "uf"];
    const rows = STATE.tableRows.map((r) => headers.map((h) => csvEscape(r[h])));
    const csv = [headers.join(";"), ...rows.map((r) => r.join(";"))].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "veiculos_midia_filtrado.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function normalizeKeys(row) {
    const out = {};
    Object.entries(row).forEach(([k, v]) => {
      const base = normalizeKeyLabel(k);
      out[base] = typeof v === "string" ? v.trim() : v;

      const coop = coopKeyFromLabel(k);
      if (coop) out[`coop_${coop}`] = v;
    });
    return out;
  }

  function normalizeKeyLabel(k) {
    const map = {
      "município sede": "municipio_sede",
      "nome da emissora": "nome_da_emissora",
      "cidades de cobertura (sinal principal)": "cidades_de_cobertura_sinal_principal",
      "população sede (hab.)": "populacao_sede_hab",
      "população total de todas cidades com sinal": "populacao_total_de_todas_cidades_com_sinal",
      "qtde associados no município sede": "qtde_associados_no_municipio_sede",
      "qtde associados município sede": "qtde_associados_municipio_sede"
    };
    const key = cleanText(k).toLowerCase();
    return map[key] || normalize(key).replace(/\s+/g, "_");
  }

  function coopKeyFromLabel(label) {
    const txt = cleanText(label).replace(/\s+/g, "");
    const n = txt.replace(/\./g, "");
    if (COOPS.includes(n)) return n;
    return null;
  }

  function activeCoopsFromRow(nRow) {
    return COOPS.filter((c) => isX(nRow[`coop_${c}`]));
  }

  function parseCoverageCities(raw, defaultUf) {
    if (!raw) return [];
    return raw
      .split(",")
      .map((chunk) => chunk.trim())
      .filter(Boolean)
      .map((chunk) => {
        const match = chunk.match(/^(.*?)\s*-\s*([A-Za-z]{2})$/);
        if (match) return { municipio: cleanCity(match[1]), uf: cleanText(match[2]).toUpperCase() };
        return { municipio: cleanCity(chunk), uf: defaultUf || "" };
      });
  }

  function dedupeMunicipios(items) {
    const map = new Map();
    items.forEach((m) => {
      const key = `${normalize(m.municipio)}|${m.uf}`;
      if (!key || key.startsWith("|")) return;
      if (!map.has(key)) map.set(key, { ...m, cooperativas: [...new Set(m.cooperativas || [])] });
      else {
        const cur = map.get(key);
        cur.cooperativas = [...new Set([...(cur.cooperativas || []), ...(m.cooperativas || [])])];
        cur.populacao = cur.populacao ?? m.populacao;
        cur.associados = cur.associados ?? m.associados;
        if (m.papel === "sede") cur.papel = "sede";
      }
    });
    return [...map.values()];
  }

  function setStatus(message, cls) {
    dom.statusBar.className = `status ${cls || ""}`.trim();
    dom.statusBar.textContent = message;
  }

  function getFeatureMunicipio(feature) {
    return feature.properties.name || feature.properties.nome || feature.properties.NOME || "";
  }

  function cleanText(v) {
    return String(v || "").replace(/\s+/g, " ").trim();
  }

  function cleanCity(v) {
    return cleanText(v).replace(/\s*\([^)]*\)\s*/g, "").trim();
  }

  function normalize(v) {
    return cleanText(v)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function toNumber(v) {
    if (v === null || v === undefined || v === "") return null;
    const n = Number(String(v).replace(/\./g, "").replace(/,/g, ".").replace(/[^\d.-]/g, ""));
    return Number.isFinite(n) ? n : null;
  }

  function isX(v) {
    return normalize(String(v || "")) === "x";
  }

  function fmt(v) {
    if (v === null || v === undefined || v === "") return "-";
    if (typeof v === "number") return new Intl.NumberFormat("pt-BR").format(v);
    return String(v);
  }

  function csvEscape(v) {
    const text = String(v ?? "");
    return `"${text.replaceAll('"', '""')}"`;
  }

  function sortCompare(a, b, asc) {
    const dir = asc ? 1 : -1;
    const an = typeof a === "number" ? a : Number(a);
    const bn = typeof b === "number" ? b : Number(b);
    if (Number.isFinite(an) && Number.isFinite(bn)) return (an - bn) * dir;
    return String(a ?? "").localeCompare(String(b ?? ""), "pt-BR") * dir;
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
})();
