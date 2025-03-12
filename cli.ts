import fs from "fs";

import {
  generateFullItinerary,
  generateIndividualItineraries,
  generateSummaryItinerary,
  loadData,
  type LoadDataOptions,
  type Data,
} from "./generate-documents.ts";
import { Packer, type Document } from "docx";

function loadDataFromFile(inputFileName: string, options: LoadDataOptions) {
  let content = fs.readFileSync(inputFileName);
  return loadData(content, options);
}

async function generateDocumentToFile(fn: () => Document, fileName: string) {
  let doc = fn();
  let content = await Packer.toBuffer(doc);
  fs.writeFileSync(fileName, content);
}

function generateFullItineraryToFile(data: Data, outputFileName: string) {
  return generateDocumentToFile(
    () => generateFullItinerary(data),
    outputFileName
  );
}

function generateIndividualItinerariesToFile(
  data: Data,
  outputFileName: string
) {
  return generateDocumentToFile(
    () => generateIndividualItineraries(data),
    outputFileName
  );
}

function generateSummaryItineraryToFile(data: Data, outputFileName: string) {
  return generateDocumentToFile(
    () => generateSummaryItinerary(data),
    outputFileName
  );
}

async function main() {
  let args = process.argv.slice(2);

  let flagsOn = true;
  let shouldGenerateFullItinerary = false;
  let fullItineraryOutputFile = null;
  let shouldGenerateIndividualItinerary = false;
  let individualItineraryOutputFile = null;
  let shouldGenerateSummaryItinerary = false;
  let summaryItineraryOutputFile = null;
  let filename = null;

  for (let argIndex = 0; argIndex < args.length; argIndex++) {
    let arg = args[argIndex];

    if (flagsOn && arg.startsWith("-")) {
      let argumentHandled = false;
      if (arg === "--") {
        argumentHandled = true;
        flagsOn = false;
      }

      if (arg === "--full") {
        argumentHandled = true;
        shouldGenerateFullItinerary = true;
      }

      if (arg === "--full-out") {
        argumentHandled = true;
        shouldGenerateFullItinerary = true;
        argIndex++;
        if (argIndex >= args.length)
          throw new Error("--full-out requires file name");
        fullItineraryOutputFile = args[argIndex];
      }

      if (arg === "--individual") {
        argumentHandled = true;
        shouldGenerateIndividualItinerary = true;
      }

      if (arg === "--individual-out") {
        argumentHandled = true;
        shouldGenerateIndividualItinerary = true;
        argIndex++;
        if (argIndex >= args.length)
          throw new Error("--individual-out requires file name");
        individualItineraryOutputFile = args[argIndex];
      }

      if (arg === "--summary") {
        argumentHandled = true;
        shouldGenerateSummaryItinerary = true;
      }

      if (arg === "--summary-out") {
        argumentHandled = true;
        shouldGenerateSummaryItinerary = true;
        argIndex++;
        if (argIndex >= args.length)
          throw new Error("--summary-out requires file name");
        summaryItineraryOutputFile = args[argIndex];
      }

      if (!argumentHandled) throw new Error("unknown option: " + arg);
    } else {
      if (filename !== null) {
        throw new Error("only supports one file");
      }

      filename = arg;
    }
  }

  if (!filename) {
    throw new Error("missing required filename");
  }

  // let data = loadData(filename, {
  //   teamRoleSource: { type: "participantsTable" },
  //   meetingRange: 2,
  // });
  let data = loadDataFromFile(filename, {
    teamRoleSource: { type: "meetingsTable", nameRow: 0, headerRow: 2 },
    meetingRange: 2,
  });

  if (
    !shouldGenerateFullItinerary &&
    !shouldGenerateIndividualItinerary &&
    !shouldGenerateSummaryItinerary
  ) {
    throw new Error(
      "no outputs selected, use --full or --individual or --summary"
    );
  }

  let tasks: Promise<void>[] = [];

  if (shouldGenerateFullItinerary) {
    fullItineraryOutputFile ??= "full-itinerary.docx";
    tasks.push(generateFullItineraryToFile(data, fullItineraryOutputFile));
  }

  if (shouldGenerateIndividualItinerary) {
    individualItineraryOutputFile ??= "individual-itineraries.docx";
    tasks.push(
      generateIndividualItinerariesToFile(data, individualItineraryOutputFile)
    );
  }

  if (shouldGenerateSummaryItinerary) {
    summaryItineraryOutputFile ??= "summary-itinerary.docx";
    tasks.push(
      generateSummaryItineraryToFile(data, summaryItineraryOutputFile)
    );
  }

  let results = await Promise.allSettled(tasks);
  for (let result of results) {
    if (result.status === "rejected") {
      throw result.reason;
    }
  }
}

await main();
