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

let vocabData = [];

const ACCEPTED_FILENAME = "vokabeln.txt";
const MAX_ENTRIES = 50;

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
  resetFeedback();
  resultMessage.classList.remove("visible", "success", "error");
  resultMessage.textContent = "";

  if (!file) {
    return;
  }

  if (file.name !== ACCEPTED_FILENAME) {
    showFeedback(uploadMessage, "error", `The file must be named <strong>${ACCEPTED_FILENAME}</strong>. Please rename your file and try again.`);
    resetFileInput();
    hideTrainer();
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    const text = reader.result;
    const { entries, skipped } = parseVocabulary(text);

    if (!entries.length) {
      showFeedback(uploadMessage, "error", "No valid vocabulary entries were found. Please check the file format and try again.");
      hideTrainer();
      return;
    }

    if (entries.length > MAX_ENTRIES) {
      showFeedback(uploadMessage, "error", `The file can contain at most ${MAX_ENTRIES} entries. Found ${entries.length}. Please reduce the list and upload again.`);
      hideTrainer();
      return;
    }

    vocabData = entries;
    renderVocabulary(entries);
    showFeedback(
      uploadMessage,
      "success",
      skipped
        ? `File loaded successfully. Skipped ${skipped} line(s) with formatting issues.`
        : "File loaded successfully. Have fun practising!"
    );
    trainerSection.classList.remove("hidden");
    toggleTrainerButtons(true);
    clearAnswers();
  };

  reader.onerror = () => {
    showFeedback(uploadMessage, "error", "The file could not be read. Please try again.");
    hideTrainer();
  };

  reader.readAsText(file, "UTF-8");
}

function parseVocabulary(text) {
  const lines = text.split(/\r?\n/);
  const entries = [];
  let skipped = 0;

  for (const [index, rawLine] of lines.entries()) {
    const line = rawLine.trim();
    if (!line) {
      continue;
    }

    const parts = line.split(";");
    if (parts.length !== 2) {
      skipped += 1;
      continue;
    }

    const german = parts[0].trim();
    const english = parts[1].trim();

    if (!german || !english) {
      skipped += 1;
      continue;
    }

    entries.push({
      german,
      english,
      line: index + 1,
    });
  }

  return { entries, skipped };
}

function renderVocabulary(entries) {
  vocabList.innerHTML = "";

  entries.forEach((entry, index) => {
    const row = document.createElement("div");
    row.className = "vocab-item";

    const germanLabel = document.createElement("span");
    germanLabel.textContent = entry.german;

    const englishInput = document.createElement("input");
    englishInput.type = "text";
    englishInput.name = `vocab-${index}`;
    englishInput.placeholder = "Enter target word";
    englishInput.autocomplete = "off";
    englishInput.dataset.answer = entry.english;
    englishInput.dataset.german = entry.german;

    englishInput.addEventListener("input", () => {
      englishInput.classList.remove("correct", "incorrect");
      resultMessage.classList.remove("visible", "success", "error");
      resultMessage.textContent = "";
    });

    row.appendChild(germanLabel);
    row.appendChild(englishInput);
    vocabList.appendChild(row);
  });
}

function compareAnswers() {
  const inputs = vocabList.querySelectorAll("input");
  if (!inputs.length) {
    return;
  }

  let correctCount = 0;
  const incorrectWords = [];

  inputs.forEach((input) => {
    const userAnswer = input.value.trim();
    const solution = input.dataset.answer.trim();

    if (!userAnswer) {
      input.classList.remove("correct");
      input.classList.add("incorrect");
      incorrectWords.push(`${input.dataset.german} (missing target word)`);
      return;
    }

    if (normalize(userAnswer) === normalize(solution)) {
      input.classList.remove("incorrect");
      input.classList.add("correct");
      correctCount += 1;
    } else {
      input.classList.remove("correct");
      input.classList.add("incorrect");
      incorrectWords.push(`${input.dataset.german} -> ${solution}`);
    }
  });

  const total = inputs.length;
  const allCorrect = correctCount === total;

  showFeedback(
    resultMessage,
    allCorrect ? "success" : "error",
    allCorrect
      ? "Great job - every target word is correct!"
      : buildErrorMessage(correctCount, total, incorrectWords)
  );
}

function buildErrorMessage(correctCount, total, incorrectWords) {
  const mistakes =
    incorrectWords.length > 0
      ? `<br><span class="details">Needs review: ${incorrectWords.join(
          ", "
        )}</span>`
      : "";

  return `You got ${correctCount} of ${total} answers right.${mistakes}`;
}

function clearAnswers(showConfirmation = false) {
  const inputs = vocabList.querySelectorAll("input");
  inputs.forEach((input) => {
    input.value = "";
    input.classList.remove("correct", "incorrect");
  });

  if (showConfirmation) {
    showFeedback(resultMessage, "success", "All answers have been cleared.");
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
  element.classList.remove("success", "error", "visible");
  element.innerHTML = message;
  element.classList.add(type, "visible");
}

function resetFeedback() {
  uploadMessage.classList.remove("success", "error", "visible");
  uploadMessage.textContent = "";
}

function resetFileInput() {
  fileInput.value = "";
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
