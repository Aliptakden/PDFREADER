import * as pdfjsLib from "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.min.mjs";

const form = document.getElementById("upload-form");
const fileInput = document.getElementById("pdf-file");
const filename = document.getElementById("filename");
const statusBox = document.getElementById("status");
const submitButton = document.getElementById("submit-button");
const dropzone = document.getElementById("dropzone");
const dpiInput = document.getElementById("dpi");
const progressFill = document.getElementById("progress-fill");
const progressText = document.getElementById("progress-text");
const summaryPanel = document.getElementById("summary-panel");
const summaryCopy = document.getElementById("summary-copy");
const summaryBody = document.getElementById("summary-body");

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.worker.min.mjs";

const partNumberPattern = /^[A-Z0-9]+(?:-[A-Z0-9]+)+[A-Z0-9]*$/;
const partNumberSearch = /[A-Z][A-Z0-9]*(?:-[A-Z0-9]+)+[A-Z0-9]*/;

const setStatus = (message, kind = "") => {
  statusBox.textContent = message;
  statusBox.className = `status ${kind}`.trim();
};

const setProgress = (value, message = "") => {
  progressFill.style.width = `${Math.max(0, Math.min(100, value))}%`;
  progressText.textContent = message;
};

const updateFilename = () => {
  filename.textContent = fileInput.files.length ? fileInput.files[0].name : "No file selected";
};

const cleanCandidate = (value) =>
  value
    .toUpperCase()
    .replace(/[^A-Z0-9-]/g, "")
    .replace(/(?<=-)[IL](?=\d|$)/g, "1")
    .replace(/(?<=\d)[IL](?=-)/g, "1");

const normalizePartNumber = (remainder) => {
  const tokens = remainder.split(/\s+/).filter(Boolean);
  if (!tokens.length) {
    return "";
  }

  const candidates = [
    cleanCandidate(tokens[0]),
    cleanCandidate(tokens.slice(0, 2).join("")),
    cleanCandidate(tokens.slice(0, 3).join("")),
    cleanCandidate(tokens.join("")),
  ];

  for (const candidate of candidates) {
    if (partNumberPattern.test(candidate)) {
      return candidate;
    }
    const match = candidate.match(partNumberSearch);
    if (match) {
      return match[0];
    }
  }

  return cleanCandidate(tokens[0]);
};

const normalizeQty = (token) => {
  const normalized = token.toUpperCase().replaceAll("I", "1").replaceAll("L", "1").replaceAll("|", "1");
  const digits = normalized.replace(/\D/g, "");
  if (!digits) {
    throw new Error("Invalid quantity");
  }
  return Number.parseInt(digits, 10);
};

const parseShipmentLine = (line, pageNumber) => {
  const tokens = line.split(/\s+/).filter(Boolean);
  if (tokens.length < 5 || !/^[0-9Il|]+$/.test(tokens[0])) {
    return null;
  }

  let quantity;
  try {
    quantity = normalizeQty(tokens[0]);
  } catch {
    return null;
  }

  let index = 1;
  const poParts = [];
  while (index < tokens.length && /^\d+$/.test(tokens[index])) {
    poParts.push(tokens[index]);
    const currentPo = poParts.join("");
    if (currentPo.startsWith("700") && currentPo.length >= 8) {
      break;
    }
    index += 1;
  }

  const poNumber = poParts.join("");
  if (!poNumber.startsWith("700") || poNumber.length < 8) {
    return null;
  }

  index += 1;
  if (index < tokens.length && tokens[index].toUpperCase() === "SO") {
    index += 1;
  }

  if (index >= tokens.length || !/^\d+$/.test(tokens[index])) {
    return null;
  }

  const packingSlip = tokens[index];
  index += 1;

  const remainder = tokens.slice(index).join(" ");
  const partNumber = normalizePartNumber(remainder);
  if (!partNumber) {
    return null;
  }

  return {
    pageNumber,
    quantity,
    poNumber,
    packingSlip,
    partNumber,
    rawLine: line,
  };
};

const extractRowsFromText = (text, pageNumber) => {
  const rows = [];
  for (const rawLine of text.split(/\r?\n/)) {
    const cleaned = rawLine.replace(/\s+/g, " ").trim();
    if (!cleaned || cleaned.includes("Qty Shipped") || cleaned.includes("P.O. Number")) {
      continue;
    }
    const row = parseShipmentLine(cleaned, pageNumber);
    if (row) {
      rows.push(row);
    }
  }
  return rows;
};

const aggregatePartTotals = (rows) => {
  const totals = new Map();
  for (const row of rows) {
    totals.set(row.partNumber, (totals.get(row.partNumber) || 0) + row.quantity);
  }
  return totals;
};

const aggregatePoPartTotals = (rows) => {
  const totals = new Map();
  for (const row of rows) {
    const key = `${row.poNumber}||${row.partNumber}`;
    totals.set(key, (totals.get(key) || 0) + row.quantity);
  }
  return totals;
};

const buildCsv = (partTotals) => {
  const lines = ["part_number,total_qty_shipped"];
  [...partTotals.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([partNumber, total]) => {
      lines.push(`${partNumber},${total}`);
    });
  return lines.join("\n");
};

const buildExcelBlob = (rows, partTotals, poPartTotals) => {
  const partTotalsData = [...partTotals.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([partNumber, totalQtyShipped]) => ({
      part_number: partNumber,
      total_qty_shipped: totalQtyShipped,
    }));

  const poPartTotalsData = [...poPartTotals.entries()]
    .map(([key, totalQtyShipped]) => {
      const [poNumber, partNumber] = key.split("||");
      return { po_number: poNumber, part_number: partNumber, total_qty_shipped: totalQtyShipped };
    })
    .sort((a, b) => a.po_number.localeCompare(b.po_number) || a.part_number.localeCompare(b.part_number));

  const detailData = rows.map((row) => ({
    page_number: row.pageNumber,
    po_number: row.poNumber,
    packing_slip: row.packingSlip,
    part_number: row.partNumber,
    quantity: row.quantity,
    raw_line: row.rawLine,
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(partTotalsData), "Part Totals");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(poPartTotalsData), "PO Part Totals");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailData), "Detail");
  const arrayBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new Blob([arrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};

const renderSummary = (partTotals, rowCount) => {
  const sorted = [...partTotals.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  summaryBody.innerHTML = "";
  sorted.slice(0, 15).forEach(([partNumber, total]) => {
    const tr = document.createElement("tr");
    const partTd = document.createElement("td");
    const totalTd = document.createElement("td");
    partTd.textContent = partNumber;
    totalTd.textContent = String(total);
    tr.append(partTd, totalTd);
    summaryBody.appendChild(tr);
  });
  summaryCopy.textContent = `${rowCount} shipment rows detected across ${partTotals.size} unique part numbers.`;
  summaryPanel.classList.remove("hidden");
};

const saveBlob = async (blob, suggestedName) => {
  if ("showSaveFilePicker" in window) {
    const handle = await window.showSaveFilePicker({
      suggestedName,
      types: [
        {
          description: "ZIP archive",
          accept: { "application/zip": [".zip"] },
        },
      ],
    });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = suggestedName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const createZipBlob = async (baseName, csvText, excelBlob) => {
  const zip = new JSZip();
  zip.file(`${baseName}_part_totals.csv`, csvText);
  zip.file(`${baseName}_part_totals.xlsx`, excelBlob);
  return zip.generateAsync({ type: "blob" });
};

const pageToCanvas = async (page, dpi) => {
  const viewport = page.getViewport({ scale: dpi / 72 });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { willReadFrequently: true });
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  await page.render({ canvasContext: context, viewport }).promise;
  return canvas;
};

const runOcr = async (file, dpi) => {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const worker = await Tesseract.createWorker("eng", 1, {
    logger: (message) => {
      if (message.status) {
        setProgress(progressFill.dataset.current || 0, message.status);
      }
    },
  });

  const rows = [];
  try {
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const percent = Math.round(((pageNumber - 1) / pdf.numPages) * 100);
      progressFill.dataset.current = String(percent);
      setProgress(percent, `Rendering page ${pageNumber} of ${pdf.numPages}`);
      const page = await pdf.getPage(pageNumber);
      const canvas = await pageToCanvas(page, dpi);
      progressFill.dataset.current = String(Math.round((pageNumber / pdf.numPages) * 100));
      setProgress(Math.round((pageNumber / pdf.numPages) * 100), `OCR page ${pageNumber} of ${pdf.numPages}`);
      const result = await worker.recognize(canvas);
      rows.push(...extractRowsFromText(result.data.text, pageNumber));
    }
  } finally {
    await worker.terminate();
  }

  return rows;
};

["dragenter", "dragover"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (event) => {
  const files = event.dataTransfer.files;
  if (files?.length) {
    fileInput.files = files;
    updateFilename();
  }
});

fileInput.addEventListener("change", updateFilename);

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const file = fileInput.files[0];
  if (!file) {
    setStatus("Choose a PDF file first.", "error");
    return;
  }

  const dpi = Number.parseInt(dpiInput.value, 10) || 220;
  const baseName = file.name.replace(/\.pdf$/i, "").replace(/[^\w.-]+/g, "_") || "report";

  submitButton.disabled = true;
  summaryPanel.classList.add("hidden");
  setProgress(0, "");
  setStatus("Processing PDF in your browser. Larger files can take a few minutes.");

  try {
    const rows = await runOcr(file, dpi);
    if (!rows.length) {
      throw new Error("No shipment rows were detected in that PDF.");
    }

    const partTotals = aggregatePartTotals(rows);
    const poPartTotals = aggregatePoPartTotals(rows);
    const csvText = buildCsv(partTotals);
    const excelBlob = buildExcelBlob(rows, partTotals, poPartTotals);
    const zipBlob = await createZipBlob(baseName, csvText, excelBlob);

    renderSummary(partTotals, rows.length);
    await saveBlob(zipBlob, `${baseName}_outputs.zip`);
    setProgress(100, "Export ready.");
    setStatus("Finished. The ZIP contains both the CSV and Excel workbook.", "success");
  } catch (error) {
    const message = error?.message || "Something went wrong while processing the PDF.";
    setStatus(message, "error");
    setProgress(0, "");
  } finally {
    submitButton.disabled = false;
  }
});
