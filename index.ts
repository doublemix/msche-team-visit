import { Packer } from "docx";
import { loadData, generateFullItinerary } from "./generate-documents.ts";

let inputFileInput: HTMLInputElement = document.getElementById(
  "input"
)! as HTMLInputElement;

let generateFullItineraryButton: HTMLButtonElement = document.getElementById(
  "generate-full-itinerary"
)! as HTMLButtonElement;

generateFullItineraryButton?.addEventListener("click", function () {
  let file = inputFileInput.files?.[0];
  if (file) {
    let reader = new FileReader();
    reader.onload = async function (event) {
      let data = event.target?.result as ArrayBuffer;
      let itinerary = loadData(data, {
        teamRoleSource: { type: "meetingsTable", nameRow: 0, headerRow: 2 },
        meetingRange: 2,
      });
      let fullItinerary = generateFullItinerary(itinerary);
      let content = await Packer.toArrayBuffer(fullItinerary);
      let blob = new Blob([content], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      let link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "full-itinerary.docx";
      link.click();
      URL.revokeObjectURL(link.href);
    };
    reader.readAsArrayBuffer(file);
  }
});
