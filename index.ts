import { Document as DocxDocument, Packer } from "docx";
import {
  loadData,
  generateFullItinerary,
  type Data,
  generateIndividualItineraries,
  generateSummaryItinerary,
} from "./generate-documents.ts";

class Observable<T> {
  value: T;
  listeners: { callback: (value: T) => void }[];

  constructor(initialValue: T) {
    this.value = initialValue;
    this.listeners = [];
  }

  with(callback: (value: T) => void) {
    this.listeners.push({ callback });
    callback(this.value);
  }

  get() {
    return this.value;
  }

  update(newValue: T) {
    this.value = newValue;

    this.notifyListeners();
  }

  private notifyListeners() {
    for (let listener of this.listeners) {
      let { callback } = listener;
      callback(this.value);
    }
  }
}

let allowedFileTypes = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
];
let dropZoneEl = document.getElementById("drop-zone")! as HTMLDivElement;

function isValidFileDrop(ev: DragEvent) {
  if (!ev.dataTransfer) return false;
  if (ev.dataTransfer.items.length !== 1) return false;
  let item = ev.dataTransfer.items[0];
  if (item.kind !== "file") return false;
  if (!allowedFileTypes.includes(item.type)) return false;
  return true;
}
dropZoneEl.addEventListener("dragenter", (ev) => {
  if (isValidFileDrop(ev)) {
    dropZoneEl.classList.add("active");
  } else {
    dropZoneEl.classList.add("invalid");
  }
});

dropZoneEl.addEventListener("dragleave", (ev) => {
  dropZoneEl.classList.remove("active");
  dropZoneEl.classList.remove("invalid");
});

dropZoneEl.addEventListener("dragover", (ev) => {
  ev.preventDefault();
});

dropZoneEl.addEventListener("drop", async (ev) => {
  dropZoneEl.classList.remove("active");
  dropZoneEl.classList.remove("invalid");
  ev.preventDefault();

  if (isValidFileDrop(ev)) {
    inputFileInput.files = ev.dataTransfer!.files;
    loadDataFromFileInput();
  }
});

let loadedDataRef = new Observable<Data | null>(null);

let inputFileInput = document.getElementById("file-input")! as HTMLInputElement;

inputFileInput.addEventListener("change", function () {
  loadDataFromFileInput();
});

function loadDataFromFileInput() {
  let file = inputFileInput.files?.[0];
  if (file) {
    loadDataFromFile(file);
  }
}

function loadDataFromFile(file: File) {
  let reader = new FileReader();
  reader.onload = function (event) {
    let fileData = event.target?.result as ArrayBuffer;
    let loadedData = loadData(fileData, {
      teamRoleSource: { type: "meetingsTable", nameRow: 0, headerRow: 2 },
      meetingRange: 2,
    });
    loadedDataRef.update(loadedData);
  };
  reader.readAsArrayBuffer(file);
}

let selectFileButton = document.getElementById(
  "select-file-button"
)! as HTMLButtonElement;

selectFileButton.addEventListener("click", () => {
  inputFileInput.click();
});

let generateFullItineraryButton = document.getElementById(
  "generate-full-itinerary"
)! as HTMLButtonElement;

let generateIndividualItinerariesButton = document.getElementById(
  "generate-individual-itineraries"
)! as HTMLButtonElement;

let generateSummaryItineraryButton = document.getElementById(
  "generate-summary-itinerary"
)! as HTMLButtonElement;

setupButton(
  generateFullItineraryButton,
  generateFullItinerary,
  "full-itinerary.docx"
);
setupButton(
  generateIndividualItinerariesButton,
  generateIndividualItineraries,
  "team-member-itineraries.docx"
);
setupButton(
  generateSummaryItineraryButton,
  generateSummaryItinerary,
  "summary-itinerary.docx"
);

function setupButton(
  button: HTMLButtonElement,
  documentGenerator: (data: Data) => DocxDocument,
  suggestedName: string
) {
  button.addEventListener("click", async function () {
    let loadedData = loadedDataRef.get();
    if (loadedData) {
      let docxDocument = documentGenerator(loadedData);
      let content = await Packer.toArrayBuffer(docxDocument);
      let blob = new Blob([content], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      let link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = suggestedName;
      link.click();
      URL.revokeObjectURL(link.href);
    }
  });
  loadedDataRef.with((loadedData) => {
    button.disabled = !loadedData;
  });
}
