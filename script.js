const fileInput = document.getElementById("fileInput");
const uploadArea = document.getElementById("uploadArea");
const uploadMessage = document.getElementById("uploadMessage");
const trainerSection = document.getElementById("trainerSection");
const vocabList = document.getElementById("vocabList");
const checkBtn = document.getElementById("checkBtn");
const clearBtn = document.getElementById("clearBtn");
const resultMessage = document.getElementById("resultMessage");
const imprintToggle = document.getElementById("imprintToggle");
const imprintPanel = document.getElementById("imprintPanel");

window.__debugLogs = [];
function debugLog(...args) {
  window.__debugLogs.push(args);
  if (typeof console !== "undefined" && typeof console.log === "function") {
    try {
      // eslint-disable-next-line no-console
      console.log("[debug]", ...args);
    } catch (error) {
      // ignore console issues
    }
  }
}

window.addEventListener("error", (event) => {
  debugLog("window.error", event.message || "unknown");
});

let vocabData = [];

const REQUIRED_EXCEL_PREFIX = "voc";
const MAX_ENTRIES = 20;
const textDecoder = new TextDecoder("utf-8");
const xmlParser = new DOMParser();

const LENGTH_BASES = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67,
  83, 99, 115, 131, 163, 195, 227, 258,
];
const LENGTH_EXTRA_BITS = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5,
  5, 5, 0,
];
const DIST_BASES = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769,
  1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DIST_EXTRA_BITS = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11,
  11, 12, 12, 13, 13,
];
const CODE_LENGTH_ORDER = [
  16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
];

const fixedHuffmanTrees = createFixedHuffmanTrees();

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  processFile(file);
});

uploadArea.addEventListener("dragover", (event) => {
  event.preventDefault();
  uploadArea.classList.add("dragover");
});

uploadArea.addEventListener("dragleave", () => {
  uploadArea.classList.remove("dragover");
});

uploadArea.addEventListener("drop", (event) => {
  event.preventDefault();
  uploadArea.classList.remove("dragover");

  const file = event.dataTransfer.files[0];
  if (!file) {
    return;
  }

  // Reflect dropped file in the hidden input for accessibility.
  if (typeof DataTransfer !== "undefined") {
    const transfer = new DataTransfer();
    transfer.items.add(file);
    fileInput.files = transfer.files;
  }

  processFile(file);
});

checkBtn.addEventListener("click", () => compareAnswers());
clearBtn.addEventListener("click", () => clearAnswers(true));
imprintToggle.addEventListener("click", () => toggleImprint());

function processFile(file) {
  debugLog("processFile:start", file && file.name);
  resetFeedback();
  resultMessage.classList.remove("visible", "success", "error");
  resultMessage.textContent = "";

  if (!file) {
    debugLog("processFile:no-file");
    return;
  }

  const extension = getExtension(file.name);
  debugLog("processFile:extension", extension);

  if (extension === "xlsx") {
    if (!file.name.toLowerCase().startsWith(REQUIRED_EXCEL_PREFIX)) {
      showFeedback(
        uploadMessage,
        "error",
        `The Excel file name must begin with <strong>${REQUIRED_EXCEL_PREFIX}</strong>. Please rename it and try again.`
      );
      resetFileInput();
      hideTrainer();
      debugLog("processFile:invalid-prefix", file.name);
      return;
    }
    showFeedback(uploadMessage, "info", "Processing Excel file …");
    readExcelFile(file);
    return;
  }

  showFeedback(
    uploadMessage,
    "error",
    "Unsupported file type. Please upload an Excel workbook with a <strong>.xlsx</strong> extension whose name starts with <strong>voc</strong>."
  );
  hideTrainer();
  resetFileInput();
  debugLog("processFile:unsupported", extension);
}

function renderVocabulary(entries) {
  debugLog("renderVocabulary", entries.length);
  vocabList.innerHTML = "";

  entries.forEach((entry, index) => {
    const row = document.createElement("div");
    row.className = "vocab-item";

    const wordCell = document.createElement("div");
    wordCell.className = "vocab-cell word-cell";
    const germanLabel = document.createElement("span");
    germanLabel.className = "word";
    germanLabel.textContent = entry.german;
    wordCell.appendChild(germanLabel);

    const answerCell = document.createElement("div");
    answerCell.className = "vocab-cell answer-cell";
    const englishInput = document.createElement("input");
    englishInput.type = "text";
    englishInput.name = `vocab-${index}`;
    englishInput.placeholder = "Enter target word";
    englishInput.autocomplete = "off";
    englishInput.dataset.answer = entry.english;
    englishInput.dataset.german = entry.german;
    answerCell.appendChild(englishInput);

    const solutionCell = document.createElement("div");
    solutionCell.className = "vocab-cell solution-cell";
    const solutionDisplay = document.createElement("span");
    solutionDisplay.className = "solution";
    solutionDisplay.textContent = "—";
    solutionDisplay.dataset.answer = entry.english;
    solutionCell.appendChild(solutionDisplay);

    englishInput.addEventListener("input", () => {
      englishInput.classList.remove("correct", "incorrect");
      resultMessage.classList.remove("visible", "success", "error");
      resultMessage.textContent = "";
      const solutionElement = englishInput
        .closest(".vocab-item")
        .querySelector(".solution");
      if (solutionElement) {
        solutionElement.textContent = "—";
        solutionElement.classList.remove(
          "solution-correct",
          "solution-incorrect",
          "revealed"
        );
      }
    });

    row.appendChild(wordCell);
    row.appendChild(answerCell);
    row.appendChild(solutionCell);
    vocabList.appendChild(row);
  });
}

function compareAnswers() {
  debugLog("compareAnswers");
  const inputs = vocabList.querySelectorAll("input");
  if (!inputs.length) {
    debugLog("compareAnswers:no-inputs");
    return;
  }

  let correctCount = 0;

  inputs.forEach((input) => {
    const userAnswer = input.value.trim();
    const solution = input.dataset.answer.trim();
    const row = input.closest(".vocab-item");
    const solutionDisplay = row ? row.querySelector(".solution") : null;
    if (solutionDisplay) {
      solutionDisplay.textContent = solution;
      solutionDisplay.classList.add("revealed");
    }

    if (!userAnswer) {
      input.classList.remove("correct");
      input.classList.add("incorrect");
      if (solutionDisplay) {
        solutionDisplay.classList.remove("solution-correct");
        solutionDisplay.classList.add("solution-incorrect");
      }
      return;
    }

    if (normalize(userAnswer) === normalize(solution)) {
      input.classList.remove("incorrect");
      input.classList.add("correct");
      correctCount += 1;
      if (solutionDisplay) {
        solutionDisplay.classList.remove("solution-incorrect");
        solutionDisplay.classList.add("solution-correct");
      }
    } else {
      input.classList.remove("correct");
      input.classList.add("incorrect");
      if (solutionDisplay) {
        solutionDisplay.classList.remove("solution-correct");
        solutionDisplay.classList.add("solution-incorrect");
      }
    }
  });

  const total = inputs.length;
  const allCorrect = correctCount === total;

  showFeedback(
    resultMessage,
    allCorrect ? "success" : "error",
    `Score: ${correctCount} / ${total}`
  );
}

function clearAnswers(showConfirmation = false) {
  debugLog("clearAnswers", showConfirmation);
  const inputs = vocabList.querySelectorAll("input");
  inputs.forEach((input) => {
    input.value = "";
    input.classList.remove("correct", "incorrect");
    const solutionDisplay = input.closest(".vocab-item")?.querySelector(".solution");
    if (solutionDisplay) {
      solutionDisplay.textContent = "—";
      solutionDisplay.classList.remove("solution-correct", "solution-incorrect", "revealed");
    }
  });

  if (showConfirmation) {
    showFeedback(resultMessage, "success", "All fields cleared.");
  } else {
    resultMessage.classList.remove("visible", "success", "error");
    resultMessage.textContent = "";
  }
}

function hideTrainer() {
  trainerSection.classList.add("hidden");
  vocabList.innerHTML = "";
  toggleTrainerButtons(false);
}

function toggleTrainerButtons(enable) {
  checkBtn.disabled = !enable;
  clearBtn.disabled = !enable;
}

function normalize(text) {
  const lowered = text.toLocaleLowerCase("de-DE").normalize("NFD");
  try {
    return lowered.replace(/\p{Diacritic}/gu, "");
  } catch (error) {
    return lowered;
  }
}

function showFeedback(element, type, message) {
  debugLog("showFeedback", type, message);
  element.classList.remove("success", "error", "info", "visible");
  element.innerHTML = message;
  element.classList.add(type, "visible");
}

function resetFeedback() {
  uploadMessage.classList.remove("success", "error", "info", "visible");
  uploadMessage.textContent = "";
}

function resetFileInput() {
  fileInput.value = "";
}

async function readExcelFile(file) {
  debugLog("readExcelFile:start", file && file.name);
  if (!supportsZipDecompression()) {
    showFeedback(
      uploadMessage,
      "error",
      "Excel decoding needs a modern browser. Please update your browser and try again."
    );
    hideTrainer();
    debugLog("readExcelFile:no-zip-support");
    return;
  }

  try {
    const buffer = await fileToArrayBuffer(file);
    debugLog("readExcelFile:buffer", buffer && buffer.byteLength);
    const bytes = new Uint8Array(buffer);
    const zip = createZipReader(bytes);
    const workbookXml = await zip.getText("xl/workbook.xml");
    debugLog("readExcelFile:workbookXml", workbookXml && workbookXml.length);
    if (!workbookXml) {
      showFeedback(uploadMessage, "error", "The Excel file is missing workbook information.");
      hideTrainer();
      debugLog("readExcelFile:no-workbook");
      return;
    }

    const workbookDoc = xmlParser.parseFromString(workbookXml, "application/xml");
    const sheets = getElementsByTagNameSafe(workbookDoc, "sheet");
    debugLog("readExcelFile:sheets", sheets.length);
    const firstSheet = sheets.length ? sheets[0] : null;
    if (!firstSheet) {
      showFeedback(uploadMessage, "error", "No worksheets found inside the Excel file.");
      hideTrainer();
      debugLog("readExcelFile:no-sheet");
      return;
    }

    const relId = firstSheet.getAttribute("r:id");
    const relsXml = await zip.getText("xl/_rels/workbook.xml.rels");
    debugLog("readExcelFile:relId", relId);
    const targetPath = normalizeSheetPath(relId && relsXml ? getRelationshipTarget(relsXml, relId) : null);
    debugLog("readExcelFile:targetPath", targetPath);

    let sheetXml = await zip.getText(targetPath);
    if (!sheetXml && !targetPath.startsWith("xl/")) {
      sheetXml = await zip.getText(`xl/${targetPath}`);
    }
    if (!sheetXml) {
      showFeedback(uploadMessage, "error", "The first worksheet could not be read.");
      hideTrainer();
      debugLog("readExcelFile:no-sheet-xml");
      return;
    }

    const sharedStringsXml = await zip.getText("xl/sharedStrings.xml");
    const sharedStrings = parseSharedStrings(sharedStringsXml);
    debugLog("readExcelFile:sharedStrings", sharedStrings.length);
    const rows = extractRowsFromSheet(sheetXml, sharedStrings);
    debugLog("readExcelFile:rows", rows.length);
    const { entries, skipped } = extractEntriesFromRows(rows);
    debugLog("readExcelFile:entries", entries.length, "skipped", skipped);
    handleParsedEntries(entries, {
      skipped,
      source: "excel",
    });
  } catch (error) {
    debugLog("readExcelFile:error", error && error.message);
    console.error(error);
    showFeedback(
      uploadMessage,
      "error",
      `The Excel file could not be processed. ${error && error.message ? error.message : "Please check the format and try again."}`
    );
    hideTrainer();
  }
}

function supportsZipDecompression() {
  debugLog(
    "supportsZipDecompression",
    typeof DecompressionStream,
    typeof inflateRaw
  );
  return (
    typeof DecompressionStream !== "undefined" ||
    typeof inflateRaw === "function"
  );
}

function createZipReader(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(view);
  if (eocdOffset < 0) {
    throw new Error("Invalid ZIP structure.");
  }

  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);

  const entries = new Map();
  let pointer = centralDirectoryOffset;

  for (let index = 0; index < totalEntries; index += 1) {
    const signature = view.getUint32(pointer, true);
    if (signature !== 0x02014b50) {
      throw new Error("Unexpected central directory signature.");
    }

    const compressedSize = view.getUint32(pointer + 20, true);
    const fileNameLength = view.getUint16(pointer + 28, true);
    const extraLength = view.getUint16(pointer + 30, true);
    const commentLength = view.getUint16(pointer + 32, true);
    const localHeaderOffset = view.getUint32(pointer + 42, true);
    const nameBytes = bytes.subarray(pointer + 46, pointer + 46 + fileNameLength);
    const fileName = textDecoder.decode(nameBytes);

    pointer += 46 + fileNameLength + extraLength + commentLength;

    const localSignature = view.getUint32(localHeaderOffset, true);
    if (localSignature !== 0x04034b50) {
      continue;
    }

    const generalPurpose = view.getUint16(localHeaderOffset + 6, true);
    const compressionMethod = view.getUint16(localHeaderOffset + 8, true);
    const fileNameLenLocal = view.getUint16(localHeaderOffset + 26, true);
    const extraLenLocal = view.getUint16(localHeaderOffset + 28, true);
    const dataOffset = localHeaderOffset + 30 + fileNameLenLocal + extraLenLocal;
    const compressedData = bytes.subarray(dataOffset, dataOffset + compressedSize);

    entries.set(fileName, {
      compressionMethod,
      compressedData,
      generalPurpose,
      cache: null,
    });
  }

  return {
    async getBytes(name) {
      debugLog("zip.getBytes", name);
      const entry = entries.get(name);
      if (!entry) {
        debugLog("zip.getBytes:missing", name);
        return null;
      }
      if (entry.cache) {
        debugLog("zip.getBytes:cache-hit", name, entry.cache.length);
        return entry.cache;
      }
      const data = await decompressEntry(entry);
      entry.cache = data;
       debugLog("zip.getBytes:decoded", name, data && data.length);
      return data;
    },
    async getText(name) {
       debugLog("zip.getText", name);
      const bytesResult = await this.getBytes(name);
      return bytesResult ? textDecoder.decode(bytesResult) : null;
    },
  };
}

async function decompressEntry(entry) {
  debugLog("decompressEntry", entry.compressionMethod);
  if (entry.compressionMethod === 0) {
    debugLog("decompressEntry:stored", entry.compressedData.length);
    return entry.compressedData.slice();
  }

  if (entry.compressionMethod === 8) {
    try {
      debugLog("decompressEntry:fallback-inflate:start");
      const result = inflateRaw(entry.compressedData);
      debugLog("decompressEntry:fallback-inflate:done", result.length);
      return result;
    } catch (error) {
      debugLog("decompressEntry:fallback-inflate:error", error && error.message);
      throw error;
    }
  }

  throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
}

function findEndOfCentralDirectory(view) {
  for (let offset = view.byteLength - 22; offset >= 0; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) {
      return offset;
    }
  }
  return -1;
}

function normalizeSheetPath(target) {
  if (!target) {
    return "xl/worksheets/sheet1.xml";
  }

  const raw = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
  const segments = raw.split("/");
  const stack = [];

  for (const segment of segments) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      stack.pop();
      continue;
    }
    stack.push(segment);
  }

  return stack.join("/");
}

function getRelationshipTarget(relsXml, relId) {
  debugLog("getRelationshipTarget", relId);
  if (!relsXml) {
    debugLog("getRelationshipTarget:no-rels");
    return null;
  }

  try {
    const relsDoc = xmlParser.parseFromString(relsXml, "application/xml");
    const relationships = relsDoc.getElementsByTagName("Relationship");
    for (const rel of relationships) {
      if (rel.getAttribute("Id") === relId) {
        return rel.getAttribute("Target");
      }
    }
  } catch (error) {
    debugLog("getRelationshipTarget:error", error && error.message);
    console.error(error);
  }

  return null;
}

function parseSharedStrings(xml) {
  debugLog("parseSharedStrings", xml ? xml.length : 0);
  if (!xml) {
    return [];
  }

  const doc = xmlParser.parseFromString(xml, "application/xml");
  const sharedStrings = [];
  const items = getElementsByTagNameSafe(doc, "si");

  for (const item of items) {
    let text = "";
    const textNodes = getElementsByTagNameSafe(item, "t");
    for (const node of textNodes) {
      text += node.textContent || "";
    }
    sharedStrings.push(text.trim());
  }

  return sharedStrings;
}

function extractRowsFromSheet(sheetXml, sharedStrings) {
  debugLog(
    "extractRowsFromSheet",
    sheetXml ? sheetXml.length : 0,
    sharedStrings.length
  );
  if (!sheetXml) {
    return [];
  }

  const doc = xmlParser.parseFromString(sheetXml, "application/xml");
  const rowNodes = getElementsByTagNameSafe(doc, "row");
  const rows = [];

  for (const rowNode of rowNodes) {
    const rowData = ["", ""];
    const cellNodes = getElementsByTagNameSafe(rowNode, "c");
    let fallbackIndex = 0;

    for (const cellNode of cellNodes) {
      const ref = cellNode.getAttribute("r") || "";
      let colIndex = columnNameToIndex(ref);
      if (Number.isNaN(colIndex) || colIndex < 0) {
        colIndex = fallbackIndex;
      }
      fallbackIndex += 1;

      if (colIndex > 1) {
        continue;
      }

      const value = readCellValue(cellNode, sharedStrings);
      rowData[colIndex] = value;
    }

    if (rowData[0] || rowData[1]) {
      rows.push(rowData);
    }
  }

  return rows;
}

function columnNameToIndex(reference) {
  if (!reference) {
    return NaN;
  }

  const match = reference.match(/[A-Za-z]+/);
  if (!match) {
    return NaN;
  }

  const letters = match[0].toUpperCase();
  let index = 0;

  for (let i = 0; i < letters.length; i += 1) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }

  return index - 1;
}

function readCellValue(cellNode, sharedStrings) {
  debugLog(
    "readCellValue",
    cellNode.getAttribute("r"),
    cellNode.getAttribute("t")
  );
  const type = cellNode.getAttribute("t");

  if (type === "s") {
    const valueNode = getElementsByTagNameSafe(cellNode, "v")[0];
    if (!valueNode) {
      return "";
    }
    const sharedIndex = Number(valueNode.textContent || "0");
    return sharedStrings[sharedIndex] || "";
  }

  if (type === "inlineStr") {
    const textNodes = getElementsByTagNameSafe(cellNode, "t");
    let inline = "";
    for (const node of textNodes) {
      inline += node.textContent || "";
    }
    return inline.trim();
  }

  const valueNode = getElementsByTagNameSafe(cellNode, "v")[0];
  if (!valueNode) {
    return "";
  }

  return (valueNode.textContent || "").trim();
}

function extractEntriesFromRows(rows) {
  debugLog("extractEntriesFromRows", rows.length);
  const entries = [];
  let skipped = 0;
  let processedRows = 0;

  for (const row of rows) {
    processedRows += 1;

    if (!Array.isArray(row)) {
      skipped += 1;
      continue;
    }

    const source = toCellValue(row[0]);
    const target = toCellValue(row[1]);

    if (!source || !target) {
      skipped += 1;
      continue;
    }

    entries.push({
      german: source,
      english: target,
      line: entries.length + 1,
    });

    if (entries.length === MAX_ENTRIES) {
      const remaining = rows.length - processedRows;
      if (remaining > 0) {
        skipped += remaining;
      }
      break;
    }
  }

  return { entries, skipped };
}

function toCellValue(cell) {
  debugLog("toCellValue", cell);
  if (cell === undefined || cell === null) {
    return "";
  }

  return String(cell).trim();
}

function handleParsedEntries(entries, { skipped, source }) {
  debugLog("handleParsedEntries", entries.length, skipped, source);
  if (!entries.length) {
    showFeedback(
      uploadMessage,
      "error",
      "No valid vocabulary pairs were found in the workbook."
    );
    hideTrainer();
    return;
  }

  vocabData = entries.slice(0, MAX_ENTRIES);
  renderVocabulary(vocabData);
  showFeedback(
    uploadMessage,
    "success",
    skipped
      ? `Loaded ${vocabData.length} entries. Skipped ${skipped} row(s) that were empty, incomplete, or beyond the ${MAX_ENTRIES}-row limit.`
      : `Loaded ${vocabData.length} entries. Time to practice!`
  );
  trainerSection.classList.remove("hidden");
  toggleTrainerButtons(true);
  clearAnswers();
  resetFileInput();
}

// Allow keyboard activation of the upload area for accessibility.
uploadArea.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

function toggleImprint() {
  const isHidden = imprintPanel.hasAttribute("hidden");
  if (isHidden) {
    imprintPanel.removeAttribute("hidden");
    imprintToggle.setAttribute("aria-expanded", "true");
    imprintToggle.textContent = "Hide imprint";
    imprintPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
  } else {
    imprintPanel.setAttribute("hidden", "");
    imprintToggle.setAttribute("aria-expanded", "false");
    imprintToggle.textContent = "Show imprint";
  }
}

function getExtension(filename = "") {
  debugLog("getExtension", filename);
  const parts = filename.toLowerCase().split(".");
  return parts.length > 1 ? parts.pop() : "";
}

function inflateRaw(input) {
  let position = 0;
  let bitBuffer = 0;
  let bitLength = 0;

  const ensureBits = (count) => {
    while (bitLength < count) {
      if (position >= input.length) {
        throw new Error("Unexpected end of data while inflating.");
      }
      bitBuffer |= input[position] << bitLength;
      position += 1;
      bitLength += 8;
    }
  };

  const readBits = (count) => {
    ensureBits(count);
    const value = bitBuffer & ((1 << count) - 1);
    bitBuffer >>>= count;
    bitLength -= count;
    return value;
  };

  const alignToByte = () => {
    bitBuffer = 0;
    bitLength = 0;
  };

  const decodeSymbol = (tree) => {
    if (tree.value !== undefined) {
      return tree.value;
    }

    let node = tree;
    while (node.value === undefined) {
      const bit = readBits(1);
      node = node[bit];
      if (!node) {
        throw new Error("Invalid Huffman code encountered.");
      }
    }
    return node.value;
  };

  const output = [];
  let isFinalBlock = false;

  while (!isFinalBlock) {
    isFinalBlock = readBits(1) === 1;
    const blockType = readBits(2);

    if (blockType === 0) {
      alignToByte();
      if (position + 4 > input.length) {
        throw new Error("Stored block header incomplete.");
      }
      const len = input[position] | (input[position + 1] << 8);
      const nlen = input[position + 2] | (input[position + 3] << 8);
      position += 4;

      if ((len ^ 0xffff) !== nlen) {
        throw new Error("Stored block length check failed.");
      }

      if (position + len > input.length) {
        throw new Error("Stored block exceeds input length.");
      }

      for (let i = 0; i < len; i += 1) {
        output.push(input[position + i]);
      }
      position += len;
      continue;
    }

    if (blockType !== 1 && blockType !== 2) {
      throw new Error("Unsupported DEFLATE block type.");
    }

    let literalTree;
    let distanceTree;

    if (blockType === 1) {
      literalTree = fixedHuffmanTrees.literal;
      distanceTree = fixedHuffmanTrees.distance;
    } else {
      const hlit = readBits(5) + 257;
      const hdist = readBits(5) + 1;
      const hclen = readBits(4) + 4;

      const codeLengthCodes = new Array(19).fill(0);
      for (let index = 0; index < hclen; index += 1) {
        codeLengthCodes[CODE_LENGTH_ORDER[index]] = readBits(3);
      }

      const codeLengthTree = buildHuffmanTree(codeLengthCodes);
      const totalCodes = hlit + hdist;
      const codeLengths = [];

      while (codeLengths.length < totalCodes) {
        const symbol = decodeSymbol(codeLengthTree);

        if (symbol <= 15) {
          codeLengths.push(symbol);
          continue;
        }

        if (symbol === 16) {
          if (!codeLengths.length) {
            throw new Error("Invalid repeat code in Huffman specification.");
          }
          const repeat = readBits(2) + 3;
          const previous = codeLengths[codeLengths.length - 1];
          for (let i = 0; i < repeat; i += 1) {
            codeLengths.push(previous);
          }
          continue;
        }

        if (symbol === 17) {
          const repeat = readBits(3) + 3;
          for (let i = 0; i < repeat; i += 1) {
            codeLengths.push(0);
          }
          continue;
        }

        if (symbol === 18) {
          const repeat = readBits(7) + 11;
          for (let i = 0; i < repeat; i += 1) {
            codeLengths.push(0);
          }
          continue;
        }

        throw new Error("Invalid code length symbol.");
      }

      const literalCodeLengths = codeLengths.slice(0, hlit);
      const distanceCodeLengths = codeLengths.slice(hlit, totalCodes);

      if (distanceCodeLengths.every((len) => len === 0)) {
        distanceCodeLengths[0] = 1;
      }

      literalTree = buildHuffmanTree(literalCodeLengths);
      distanceTree = buildHuffmanTree(distanceCodeLengths);
    }

    while (true) {
      const symbol = decodeSymbol(literalTree);

      if (symbol === 256) {
        break;
      }

      if (symbol < 256) {
        output.push(symbol);
        continue;
      }

      const lengthIndex = symbol - 257;
      if (lengthIndex < 0 || lengthIndex >= LENGTH_BASES.length) {
        throw new Error("Invalid length code encountered.");
      }

      const length =
        LENGTH_BASES[lengthIndex] + readBits(LENGTH_EXTRA_BITS[lengthIndex]);
      const distanceSymbol = decodeSymbol(distanceTree);

      if (distanceSymbol < 0 || distanceSymbol >= DIST_BASES.length) {
        throw new Error("Invalid distance code encountered.");
      }

      const distance =
        DIST_BASES[distanceSymbol] + readBits(DIST_EXTRA_BITS[distanceSymbol]);
      if (distance <= 0 || distance > output.length) {
        throw new Error("Invalid distance in back-reference.");
      }

      const start = output.length - distance;
      for (let i = 0; i < length; i += 1) {
        output.push(output[start + i]);
      }
    }
  }

  return Uint8Array.from(output);
}

function fileToArrayBuffer(file) {
  if (typeof file.arrayBuffer === "function") {
    return file.arrayBuffer();
  }

  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("The Excel file could not be read."));
    reader.readAsArrayBuffer(file);
  });
}

function buildHuffmanTree(lengths) {
  let maxBits = 0;
  for (const length of lengths) {
    if (length > maxBits) {
      maxBits = length;
    }
  }

  if (maxBits === 0) {
    const valueIndex = lengths.findIndex((length) => length > 0);
    return {
      value: valueIndex >= 0 ? valueIndex : 0,
    };
  }

  const blCount = new Array(maxBits + 1).fill(0);
  lengths.forEach((length) => {
    if (length > 0) {
      blCount[length] += 1;
    }
  });

  const nextCode = new Array(maxBits + 1).fill(0);
  let code = 0;
  for (let bits = 1; bits <= maxBits; bits += 1) {
    code = (code + (blCount[bits - 1] || 0)) << 1;
    nextCode[bits] = code;
  }

  const root = {};

  lengths.forEach((length, symbol) => {
    if (length === 0) {
      return;
    }

    let node = root;
    let codeValue = nextCode[length];
    nextCode[length] += 1;

    for (let bitIndex = length - 1; bitIndex >= 0; bitIndex -= 1) {
      const bit = (codeValue >> bitIndex) & 1;
      node[bit] = node[bit] || {};
      node = node[bit];
    }

    node.value = symbol;
  });

  return root;
}

function createFixedHuffmanTrees() {
  const literalLengths = new Array(288);

  for (let i = 0; i <= 143; i += 1) {
    literalLengths[i] = 8;
  }
  for (let i = 144; i <= 255; i += 1) {
    literalLengths[i] = 9;
  }
  for (let i = 256; i <= 279; i += 1) {
    literalLengths[i] = 7;
  }
  for (let i = 280; i <= 287; i += 1) {
    literalLengths[i] = 8;
  }

  const distanceLengths = new Array(32).fill(5);

  return {
    literal: buildHuffmanTree(literalLengths),
    distance: buildHuffmanTree(distanceLengths),
  };
}

function getElementsByTagNameSafe(node, tagName) {
  if (!node || typeof node.getElementsByTagName !== "function") {
    return [];
  }

  const direct = node.getElementsByTagName(tagName);
  if (direct && direct.length) {
    return Array.from(direct);
  }

  if (typeof node.getElementsByTagNameNS === "function") {
    const namespaceList = [
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "*",
    ];

    for (const ns of namespaceList) {
      try {
        const nsNodes = node.getElementsByTagNameNS(ns, tagName);
        if (nsNodes && nsNodes.length) {
          return Array.from(nsNodes);
        }
      } catch (error) {
        // Ignore namespace lookup errors and continue.
      }
    }
  }

  return [];
}
