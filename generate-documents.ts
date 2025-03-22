import {
  utils as xlsxUtils,
  read as xlsxRead,
  type WorkBook,
  type WorkSheet,
  type CellObject,
} from "xlsx";
import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Header,
  type IParagraphOptions,
  type IRunOptions,
  type ISectionOptions,
  type ISpacingProperties,
  LevelFormat,
  Packer,
  PageNumber,
  Paragraph,
  type ParagraphChild,
  Tab,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import { parse as parseSimpleXml } from "./simple-xml.js";

export class UserError extends Error {
  constructor(message?: string) {
    super(message);
  }
}

interface Message {
  type: "log" | "warn" | "error" | "codeError";
  messageContent: string;
}
export class MessageCollector {
  messages: Message[];
  constructor() {
    this.messages = [];
  }
  codeError(err: Error) {
    this.pushMessage("codeError", "error", err.message, err);
  }
  error(messageContent: string, ...optionalParams: any[]) {
    this.pushMessage("error", "error", messageContent, ...optionalParams);
  }
  warn(messageContent: string, ...optionalParams: any[]) {
    this.pushMessage("warn", "warn", messageContent, ...optionalParams);
  }
  log(messageContent: string, ...optionalParams: any[]) {
    this.pushMessage("log", "log", messageContent, ...optionalParams);
  }
  pushMessage(
    type: Message["type"],
    logType: "log" | "warn" | "error",
    messageContent: string,
    ...optionalParams: any[]
  ) {
    this.messages.push({ type, messageContent });
    console[logType](messageContent, ...optionalParams);
  }
}

type SimpleXml = SimpleXmlNode[];
type SimpleXmlNode = SimpleXmlTextNode | SimpleXmlTagNode;
type SimpleXmlTextNode = { type: "text"; value: string };
type SimpleXmlTagNode = {
  type: "tag";
  tagName: string;
  attributes: { name: string; value: string }[];
  content: SimpleXml;
};

function formatSimpleXml(xml: SimpleXml) {
  type ActiveOptions = { bold?: true; italics?: true; underline?: true };

  let results: TextRun[] = [];
  let defaultOptions: ActiveOptions = {};

  handleContent();

  return results;

  function handleContent() {
    for (let item of xml) {
      // expect `r` tag`
      if (throwOnNonEmptyText(item)) continue;
      if (item.tagName === "r") handleRTag(item.content);
      else if (item.tagName === "t")
        generateTextNodeWithOptions(item.content, defaultOptions);
      else throw new Error("unexpected tag at top-level: " + item.tagName);
    }
  }

  function handleRTag(content: SimpleXml) {
    let textNodes: SimpleXmlTagNode[] = [];
    let runPropertiesNodes: SimpleXmlTagNode[] = [];
    for (let item of content) {
      if (throwOnNonEmptyText(item)) continue;
      if (item.tagName === "rPr") {
        runPropertiesNodes.push(item);
      } else if (item.tagName === "t") {
        textNodes.push(item);
      } else throw new Error("unexpected tag in `r` tag: " + item.tagName);
    }

    let activeOptions: ActiveOptions = {};

    for (let runPropertiesNode of runPropertiesNodes) {
      for (let item of runPropertiesNode.content) {
        if (throwOnNonEmptyText(item)) continue;
        if (item.tagName === "b") activeOptions.bold = true;
        if (item.tagName === "i") activeOptions.italics = true;
        if (item.tagName === "u") activeOptions.underline = true;
        // ignore other tagNames, we aren't concerned with them
      }
    }

    for (let textNode of textNodes) {
      generateTextNodeWithOptions(textNode.content, activeOptions);
    }
  }

  function generateTextNodeWithOptions(
    content: SimpleXml,
    activeOptions: ActiveOptions
  ) {
    for (let item of content) {
      if (item.type !== "text") throw new Error("only expecting text content");
      results.push(
        new TextRun({
          bold: activeOptions.bold,
          italics: activeOptions.italics,
          underline: activeOptions.underline && {
            type: "single",
          },
          text: item.value,
        })
      );
    }
  }

  function throwOnNonEmptyText(node: SimpleXmlNode): node is SimpleXmlTextNode {
    if (node.type !== "text") return false;
    if (node.value.trim() !== "") {
      throw new Error("unexpected text");
    }
    return true;
  }
}

function _throw(err: any) {
  throw err;
}

type Mutable<T> = {
  -readonly [P in keyof T]: T[P];
};

function findSingle<T>(array: T[], predicate: (x: T) => boolean) {
  let candidates = array.filter(predicate);
  if (candidates.length !== 1) {
    throw new Error("findSingle did not find 1 result");
  }
  return candidates[0] as T;
}

type MapTuple<T extends any[], R> = {
  [K in keyof T]: R;
};

type StructuredResponse<T, R, U> = T extends []
  ? []
  : T extends [infer A, ...infer B]
  ? [
      StructuredResponse<A, R, U>,
      ...Extract<StructuredResponse<B, R, U>, any[]>
    ]
  : T extends { [key: string]: infer A }
  ? { [K in keyof T]: StructuredResponse<T[K], R, U> }
  : R;

type StructuredRequest<T> =
  | { [key: string]: StructuredRequest<T> }
  | StructuredRequest<T>[]
  | T;

function applyStructuredRequest<T extends StructuredRequest<U>, R, U>(
  structuredRequest: T,
  mapper: (arg0: U) => R
): StructuredResponse<T, R, U>;
function applyStructuredRequest(structuredRequest: any, mapper: any): any {
  if (Array.isArray(structuredRequest)) {
    return structuredRequest.map((item) =>
      applyStructuredRequest(item, mapper)
    );
  } else if (typeof structuredRequest === "object") {
    return Object.fromEntries(
      Object.entries(structuredRequest).map(([key, value]) => [
        key,
        applyStructuredRequest(value, mapper),
      ])
    );
  } else {
    return mapper(structuredRequest);
  }
}

type WorkSheetRequest =
  | string
  | {
      name: string;
      range?: number;
      additionalColumns?: (
        w: WorkSheet,
        cs: Record<string, number>
      ) => (r: number) => Record<string, any>;
    };
function getWorksheets(
  workbook: WorkBook,
  worksheetRequests: WorkSheetRequest[]
) {
  return worksheetRequests.map((req) => {
    if (typeof req === "string") {
      req = { name: req };
    }
    let worksheet = workbook.Sheets[req.name];
    let parsed = xlsxUtils.sheet_to_json<Record<string, string>>(worksheet, {
      defval: "",
      skipHidden: true,
      range: req.range,
    });
    if (req.additionalColumns) {
      let r = xlsxUtils.decode_range(worksheet["!ref"]!);
      let row = req.range ?? 0;
      let columnNames: Record<string, number> = {};
      for (let column = r.s.c; column < r.e.c; column++) {
        let cell: CellObject = worksheet[row][column];
        columnNames[cell.w!] = column;
      }

      let additionalColumns = req.additionalColumns(worksheet, columnNames);

      for (let row of parsed) {
        let additional = additionalColumns(row.__rowNum__ as unknown as number);
        Object.assign(row, additional);
      }
    }
    return parsed;
  });
}

let boolean = (input: string) => {
  return input.trim() !== "";
};

let stringToBoolean = (yesText: string, noText: string) => (input: string) => {
  if (input === yesText) return true;
  if (input === noText) return false;
  throw new UserError(
    `failed to parse boolean, expected ${yesText} or ${noText}, got ${input}`
  );
};

type MapInputMatcher = string | string[] | RegExp | true;

type MapInput<T> = [MapInputMatcher, T];
let doesMatchMapInputMapper = (
  input: string,
  matcher: MapInputMatcher
): boolean => {
  if (typeof matcher === "string") {
    return input === matcher;
  }
  if (Array.isArray(matcher)) {
    return matcher.some((potenial) => potenial === input);
  }
  if (matcher instanceof RegExp) {
    return input.match(matcher) !== null;
  }
  if (matcher === true) {
    return true;
  }
  return false;
};
let mapInput =
  <T>(mappings: Array<MapInput<T>>) =>
  (input: string) => {
    for (let [matcher, output] of mappings) {
      if (doesMatchMapInputMapper(input, matcher)) {
        return output;
      }
    }
    throw new UserError(`Unmapped value for mapInput: ${input}`);
  };

type Mapper =
  | string
  | RegExp
  | ((row: any) => any)
  | Mapper[]
  | { [key: string]: Mapper };

function isPlainObject(x: any) {
  return (
    x !== null &&
    typeof x === "object" &&
    (Object.getPrototypeOf(x) === Object.prototype ||
      Object.getPrototypeOf(x) === null)
  );
}

function mapFields<T>(data: any[], mapper: any): T[];
function mapFields<T>(data: any[], mapper: Record<string, Mapper>): T[] {
  return data.map((row) => {
    return mapField(row, mapper);
  }) as T[];
}

function mapField(data: any, mapper: Mapper): any {
  if (typeof mapper === "string") {
    if (!(mapper in data))
      throw new UserError(`Field ${mapper} not found in data`);
    let value = data[mapper];
    if (typeof value === "string") {
      return value.trim();
    }
    return value ?? "";
  } else if (mapper instanceof RegExp) {
    let candidateKeys = Object.keys(data).filter((key) => key.match(mapper));
    if (candidateKeys.length === 0)
      throw new UserError(`Field ${mapper} not found in data`);
    if (candidateKeys.length > 1)
      throw new UserError(
        `Multiple keys match ${mapper}: ${candidateKeys.join(
          ", "
        )}; use a more specific expression`
      );
    let key = candidateKeys[0];
    return mapField(data, key);
  } else if (typeof mapper === "function") {
    return mapper(data);
  } else if (Array.isArray(mapper)) {
    return mapper.reduce((value, mapValue) => {
      return mapField(value, mapValue);
    }, data);
  } else if (isPlainObject(mapper)) {
    let result: Record<any, any> = {};
    for (let [key, subMapper] of Object.entries(mapper as any)) {
      let propertyDefinition = Object.getOwnPropertyDescriptor(mapper, key)!;
      if (propertyDefinition.get || propertyDefinition.set) {
        Object.defineProperty(result, key, propertyDefinition);
      } else {
        let value = mapField(data, subMapper as any);
        if (key.startsWith("$")) {
          Object.assign(result, value);
        } else {
          result[key] = mapField(data, subMapper as any);
        }
      }
    }
    return result;
  }
  throw new Error("did not understand mapper");
}

function commaSeparatedList(input: string) {
  return input
    .split(",")
    .map((item) => item.trim())
    .filter((item) => item !== "");
}

function displayNameAndId(input: string): { displayName: string; id: string } {
  return {
    id: convertToId(input),
    displayName: input,
  };
}

function convertToId(input: string) {
  return input
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, "")
    .replace(/\s+/g, "-");
}

export type Participant = {
  prefix: string;
  firstName: string;
  lastName: string;
  title: string;
  staff: boolean;
  faculty: boolean;
  email: string;
  id: string;
  teamMemberRoles: string[];
  fullName: string;
};

export interface TimeOfDay {
  hour12: number;
  hour24: number;
  minute: number;
  amPm: "a" | "p";
}

export type ProposedMeeting = {
  date: string;
  time: string;
  startTime: TimeOfDay | null;
  endTime: TimeOfDay | null;
  meetingLocation: string;
  zoomRoomOptionType: "none" | "optional" | "primary";
  shouldShowZoomRoom: boolean;
  isZoomRoomPrimary: boolean;
  zoomRoomName: string;
  interviewAssignments: string;
  teamChair: boolean;
  standard1TeamMember: boolean;
  standard2TeamMember: boolean;
  standard3TeamMember: boolean;
  standard4TeamMember: boolean;
  standard5TeamMember: boolean;
  standard6TeamMember: boolean;
  standard7TeamMember: boolean;
  individuals: { displayName: string; id: string }[];
  hideNames: boolean;
  teamRoles: string;
  teamRolesData: SimpleXml | null;
};

export type ZoomRoom = {
  zoomRoomName: string;
  link: string;
};

function toMap<T, K>(array: T[], keySelector: (x: T) => K) {
  let map = new Map<K, T>();

  for (let item of array) {
    let key = keySelector(item);
    if (map.has(key)) {
      throw new Error("duplicate key");
    }
    map.set(key, item);
  }

  return map;
}

function iife<F extends () => any>(f: F): ReturnType<F> {
  return f();
}

let { twips, borderSize } = iife(() => {
  function createConversionObject(twipsPerUnit: number) {
    function fromInches(inches: number) {
      return fromPoints(inches * 72);
    }

    function fromPoints(points: number) {
      return fromTwips(points * 20);
    }

    function fromTwips(twips: number) {
      return Math.round(twips / twipsPerUnit);
    }

    return { fromInches, fromPoints, fromTwips };
  }

  return {
    twips: createConversionObject(1),
    borderSize: createConversionObject(20 / 8), // borders are measured in eighths of a point
  };
});

type Group<T, K extends PropertyKey> = T[] & { key: K };
type Grouped<T, K extends PropertyKey> = Group<T, K>[];

let groupBy = <T, K extends PropertyKey>(
  data: T[],
  keySelector: (item: T) => K
): Grouped<T, K> => {
  let groupMap = new Map<K, Group<T, K>>();
  let groups: Group<T, K>[] = [];

  data.forEach((item) => {
    let key = keySelector(item);
    if (!groupMap.has(key)) {
      let group = Object.assign([], { key });
      groupMap.set(key, group);
      groups.push(group);
    }
    groupMap.get(key)!.push(item);
  });

  return groups;
};

let teamMemberDefinitions = [
  { property: "teamChair" as "teamChair", value: "Team Chair" },
  { property: "standard1TeamMember" as "standard1TeamMember", value: "SI" },
  { property: "standard2TeamMember" as "standard2TeamMember", value: "SII" },
  { property: "standard3TeamMember" as "standard3TeamMember", value: "SIII" },
  { property: "standard4TeamMember" as "standard4TeamMember", value: "SIV" },
  { property: "standard5TeamMember" as "standard5TeamMember", value: "SV" },
  { property: "standard6TeamMember" as "standard6TeamMember", value: "SVI" },
  { property: "standard7TeamMember" as "standard7TeamMember", value: "SVII" },
];

let teamMemberDefinitionsByRole = toMap(teamMemberDefinitions, (x) => x.value);

export type Data = {
  proposedMeetingsData: ProposedMeeting[];
  participantListData: Participant[];
  zoomRooms: ZoomRoom[];
  zoomRoomsByName: Map<string, ZoomRoom>;
};

export type LoadDataOptions = {
  meetingRange?: number;
  teamRoleSource:
    | {
        type: "meetingsTable";
        nameRow: number;
        headerRow: number;
      }
    | {
        type: "participantsTable";
      };
};

function getColumnIndex(
  matcher: string | RegExp,
  columnMap: Record<string, number>
) {
  let keys = Object.keys(columnMap);

  let matched = keys.filter((k) => k.match(matcher));

  if (matched.length < 1) {
    throw new Error("column not found: " + matcher);
  }

  if (matched.length !== 1) {
    throw new Error("multiple columns found: " + matcher);
  }

  return columnMap[matched[0]];
}

export function loadData(
  input: Buffer | ArrayBuffer,
  opts: LoadDataOptions,
  messageCollector: MessageCollector
): Data {
  // let workbook = xlsx.readFile(filename, { dense: true });
  let workbook = xlsxRead(input, { dense: true });

  let [proposedMeetingsRawData, participantListRawData, zoomRoomsRawData] =
    getWorksheets(workbook, [
      {
        name: "Proposed Meetings-MSCHE Team",
        range: opts.meetingRange,
        additionalColumns: (w, cs) => {
          let rolesColumnIndex = getColumnIndex(/host \(h\)/i, cs);
          return (r) => {
            let cell = w[r][rolesColumnIndex];
            let parsed: SimpleXml | null = null;

            if (cell?.h) {
              try {
                parsed = parseSimpleXml(cell?.r);
              } catch (err) {
                messageCollector.warn(
                  "Error while parsing XML: " +
                    cell?.h +
                    "; " +
                    (err instanceof Error ? err.message : err),
                  err
                );
              }
            }
            return { "Formatted roles": parsed };
          };
        },
      },
      "Participant List",
      "Zoom Rooms",
    ]);
  let participantListData = mapFields<Participant>(participantListRawData, {
    prefix: "PFX",
    firstName: "First Name",
    lastName: "Last Name",
    title: "Title /Involvement",
    staff: ["Staff", boolean],
    faculty: ["Faculty", boolean],
    email: "Email",
    // id: (row) => {
    //   return convertToId(
    //     `${row["PFX"]} ${row["First Name"]} ${row["Last Name"]}`
    //   );
    // },
    get fullName() {
      return `${this.prefix} ${this.firstName} ${this.lastName}`
        .replace(/\s+/, " ")
        .trim();
    },
    get id() {
      let that = this as unknown as Participant;
      return convertToId(that.fullName);
    },
    teamMemberRoles:
      opts.teamRoleSource.type === "meetingsTable"
        ? () => []
        : opts.teamRoleSource.type === "participantsTable"
        ? ["Team Member", commaSeparatedList]
        : _throw(new Error("unexpected teamRoleSource")),
  });

  if (opts.teamRoleSource.type === "meetingsTable") {
    let meetingsWorksheet = workbook.Sheets["Proposed Meetings-MSCHE Team"];
    let meetingsWorksheetFirstRow =
      meetingsWorksheet[opts.teamRoleSource.nameRow];
    let columnHeadersRow = meetingsWorksheet[opts.teamRoleSource.headerRow];

    for (let i = 0; i < meetingsWorksheetFirstRow.length; i++) {
      let teamMemberLastNameCell = meetingsWorksheetFirstRow[i];
      if (!teamMemberLastNameCell) continue;
      if (teamMemberLastNameCell.t !== "s")
        throw new Error("expected cell type to be s");
      let teamMemberLastName = teamMemberLastNameCell.v;
      if (teamMemberLastName === "") continue;

      let columnHeaderCell = columnHeadersRow[i];
      if (columnHeaderCell.t !== "s")
        throw new Error("expected cell type to be s");
      let columnName = columnHeaderCell.v;

      let candidateTeamMembers = participantListData.filter(
        (p) => p.lastName === teamMemberLastName
      );
      if (candidateTeamMembers.length < 1)
        throw new Error(
          `no candidates for ${columnName}, ${teamMemberLastName}`
        );
      if (candidateTeamMembers.length > 1)
        throw new Error(
          `multiple candidates for ${columnName}, ${teamMemberLastName}`
        );

      let teamMember = candidateTeamMembers[0];

      teamMember.teamMemberRoles.push(columnName);
    }
  }

  function getTimeFromEnd(
    hourAsString: string,
    minuteAsString: string,
    fromEndTime: TimeOfDay
  ): TimeOfDay {
    let candidateTime = getTime(hourAsString, minuteAsString, fromEndTime.amPm);
    if (candidateTime.hour24 > fromEndTime.hour24) {
      candidateTime = getTime(hourAsString, minuteAsString, "a");
    }
    return candidateTime;
  }
  function getTime(
    hourAsString: string,
    minuteAsString: string,
    amPm: string
  ): TimeOfDay {
    let hour = parseInt(hourAsString);
    let minute = parseInt(minuteAsString);

    if (!isFinite(hour)) {
      throw new Error("invalid hour");
    }
    if (!isFinite(minute)) {
      throw new Error("invalid minute");
    }

    // convert to twenty hour as
    let isPm = amPm === "p";
    let isAm = amPm === "a";

    if (!isAm && !isPm) {
      throw new Error("invalid am/pm indicator");
    }
    let hour24 = hour === 12 ? (isAm ? 0 : 12) : isPm ? hour + 12 : hour;

    return {
      hour12: hour,
      hour24: hour24,
      amPm: isAm ? "a" : "p",
      minute,
    };
  }

  let proposedMeetingsData = mapFields<ProposedMeeting>(
    proposedMeetingsRawData.filter((m) => m.Date),
    {
      date: "Date",
      $time: [
        "Time",
        (time: string) => {
          let match;
          let startTime: TimeOfDay | null = null,
            endTime: TimeOfDay | null = null;
          if (
            (match = time.match(
              /^Up to (?<hour>[0-9]+):(?<minute>[0-9]+) (?<ampm>[ap])\.m\.$/
            ))
          ) {
            endTime = getTime(
              match.groups!.hour,
              match.groups!.minute,
              match.groups!.ampm
            );
          } else if (
            (match = time.match(
              /^(?<hour>[0-9]+):(?<minute>[0-9]+) (?<ampm>[ap])\.m\.$/
            ))
          ) {
            startTime = getTime(
              match.groups!.hour,
              match.groups!.minute,
              match.groups!.ampm
            );
          } else if (
            (match = time.match(
              /^(?<startHour>[0-9]+):(?<startMinute>[0-9]+)[-–](?<endHour>[0-9]+):(?<endMinute>[0-9]+) (?<endAmPm>[ap])\.m\.$/
            ))
          ) {
            endTime = getTime(
              match.groups!.endHour,
              match.groups!.endMinute,
              match.groups!.endAmPm
            );
            startTime = getTimeFromEnd(
              match.groups!.startHour,
              match.groups!.startMinute,
              endTime
            );
          } else {
            throw new Error(`couldn't parse time: \`${time}\``);
          }

          return { time, startTime, endTime };
        },
      ],
      meetingLocation: "Meeting Location",
      zoomRoomOptionType: [
        "Zoom Room Option",
        mapInput([
          ["Primary Room", "primary"],
          ["Yes", "optional"],
          [["No", "N/A", "", "N0"], "none"],
        ]),
      ],
      get shouldShowZoomRoom() {
        let that = this as unknown as ProposedMeeting;
        return (
          that.zoomRoomOptionType === "primary" ||
          that.zoomRoomOptionType === "optional"
        );
      },
      get isZoomRoomPrimary() {
        let that = this as unknown as ProposedMeeting;
        return that.zoomRoomOptionType === "primary";
      },
      zoomRoomName: /^Zoom Link/,
      interviewAssignments: "Interview Assignments",
      teamChair: ["Team Chair", boolean],
      standard1TeamMember: ["SI", boolean],
      standard2TeamMember: ["SII", boolean],
      standard3TeamMember: ["SIII", boolean],
      standard4TeamMember: ["SIV", boolean],
      standard5TeamMember: ["SV", boolean],
      standard6TeamMember: ["SVI", boolean],
      standard7TeamMember: ["SVII", boolean],
      individuals: [
        "Individuals",
        commaSeparatedList,
        (x: string[]) => x.map((entry: string) => displayNameAndId(entry)),
      ],
      hideNames: ["Hide Names", boolean],
      teamRoles: /host \(h\)/i,
      teamRolesData: "Formatted roles",
    }
  );

  let zoomRooms = mapFields<ZoomRoom>(zoomRoomsRawData, {
    zoomRoomName: "Zoom Room Name",
    link: "Link",
  });
  let zoomRoomsByName = toMap(zoomRooms, (zr) => zr.zoomRoomName);

  return {
    proposedMeetingsData,
    participantListData,
    zoomRooms,
    zoomRoomsByName,
  };
}

function getParticipantTeamMembers(
  participantListData: Participant[],
  meeting: ProposedMeeting
) {
  let allTeamMembers = participantListData.filter(
    (p) => p.teamMemberRoles.length >= 1
  );

  let teamMembers = participantListData.filter((p) => {
    for (let d of teamMemberDefinitions) {
      if (meeting[d.property] && p.teamMemberRoles.includes(d.value)) {
        return true;
      }
    }

    return false;
  });

  let isAllTeamMembers = teamMembers.length === allTeamMembers.length;

  return {
    isAllTeamMembers,
    teamMembers,
  };
}

export function generateFullItinerary(
  data: Data,
  messageCollector: MessageCollector
) {
  const { proposedMeetingsData, participantListData, zoomRoomsByName } = data;

  let proposedMeetingsGroupedByDate = groupBy(
    proposedMeetingsData,
    (x) => x.date
  );
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "schedule-list",
          levels: [
            {
              level: 0,
              format: LevelFormat.NONE,
              alignment: "left",
              style: {
                paragraph: {
                  leftTabStop: twips.fromInches(0.5),
                  indent: {
                    left: twips.fromInches(0.75),
                    hanging: twips.fromInches(0.5),
                  },
                  spacing: {
                    beforeAutoSpacing: true,
                    afterAutoSpacing: true,
                  },
                },
                run: {
                  size: "12pt",
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.NONE,
              alignment: "left",
              style: {
                paragraph: {
                  indent: {
                    left: twips.fromInches(1),
                    hanging: twips.fromInches(0.25),
                  },
                  spacing: {
                    beforeAutoSpacing: true,
                    afterAutoSpacing: true,
                  },
                },
                run: {
                  size: "10pt",
                },
              },
            },
            {
              level: 2,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: "left",
              style: {
                paragraph: {
                  indent: {
                    left: twips.fromInches(1.5),
                    hanging: twips.fromInches(0.25),
                  },
                  spacing: {
                    beforeAutoSpacing: true,
                    afterAutoSpacing: true,
                  },
                },
                run: {
                  size: "10pt",
                },
              },
            },
          ],
        },
      ],
    },
    styles: {
      paragraphStyles: [
        {
          id: "Date",
          name: "Date",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          paragraph: {
            spacing: {
              // before: twips.fromPoints(10),
            },
          },
          run: {
            font: "Times New Roman",
            size: "14pt",
            bold: true,
            underline: {
              type: "single",
            },
          },
        },
      ],
    },
    sections: [
      Section({
        headerTitle: "Detailed Itinerary",
        children: proposedMeetingsGroupedByDate.flatMap((group) => {
          return [
            new Paragraph({
              style: "Date",
              text: group.key,
            }),
            ...group.flatMap((meeting) => {
              return [
                new Paragraph({
                  numbering: {
                    reference: "schedule-list",
                    level: 0,
                  },
                  children: iife(() => {
                    let children: ParagraphChild[] = [];

                    let segments: {
                      run: ParagraphChild;
                      following?: ParagraphChild;
                    }[] = [
                      {
                        run: new TextRun({
                          bold: true,
                          text: meeting.interviewAssignments,
                        }),
                        following: new TextRun({
                          bold: true,
                          text: `, `,
                        }),
                      },
                    ];

                    if (meeting.time) {
                      segments.push({
                        run: new TextRun({
                          color: "C00000",
                          text: meeting.time,
                        }),
                      });
                    }

                    if (meeting.meetingLocation) {
                      segments.push({
                        run: new TextRun(meeting.meetingLocation),
                      });
                    }

                    if (meeting.shouldShowZoomRoom) {
                      segments.push({
                        run: ZoomLink(
                          meeting.zoomRoomName,
                          zoomRoomsByName,
                          messageCollector
                        ),
                      });
                    }

                    let pendingFollowing: ParagraphChild | null = null;
                    let first = true;
                    for (let { run, following } of segments) {
                      if (!first) {
                        pendingFollowing ??= new TextRun(", ");
                        children.push(pendingFollowing);
                      }
                      first = false;
                      children.push(run);
                      pendingFollowing = following ?? null;
                    }

                    return children;
                  }),
                }),

                ...iife(() => {
                  if (meeting.hideNames) return [];

                  let { teamMembers, isAllTeamMembers } =
                    getParticipantTeamMembers(participantListData, meeting);

                  if (teamMembers.length < 1) {
                    return [];
                  }

                  return [
                    new Paragraph({
                      numbering: {
                        reference: "schedule-list",
                        level: 1,
                      },
                      children: [
                        new TextRun({
                          bold: true,
                          text: `MSCHE Team Member(s)`,
                        }),
                      ],
                    }),
                    ...(isAllTeamMembers
                      ? [
                          new Paragraph({
                            numbering: {
                              reference: "schedule-list",
                              level: 2,
                            },
                            children: [new TextRun("All Team Members")],
                          }),
                        ]
                      : teamMembers.map((teamMember) => {
                          return renderParticipant(teamMember);
                        })),
                  ];
                }),
                ...iife(() => {
                  if (meeting.hideNames) return [];
                  let individuals = meeting.individuals.flatMap(
                    (individual) => {
                      let participant = participantListData.find(
                        (participant) => participant.id === individual.id
                      );
                      if (!participant) {
                        messageCollector.error(
                          `missing individual in ${meeting.interviewAssignments}: ` +
                            individual.id
                        );
                        return [];
                      }
                      return [renderParticipant(participant)];
                    }
                  );

                  if (individuals.length < 1) {
                    return [];
                  }

                  return [
                    new Paragraph({
                      numbering: {
                        reference: "schedule-list",
                        level: 1,
                      },
                      children: [
                        new TextRun({
                          bold: true,
                          text: "CU Representative(s)",
                        }),
                      ],
                    }),
                    ...individuals,
                  ];
                }),
              ];
            }),
          ];
        }),
      }),
    ],
  });

  function renderParticipant(participant: Participant) {
    let title = participant?.title ?? "";
    // TODO log error
    let hasTitle = !!title;
    return new Paragraph({
      numbering: {
        reference: "schedule-list",
        level: 2,
      },
      children: iife(() => {
        let result = [
          new TextRun({
            bold: true,
            text: `${participant.fullName}${hasTitle ? ", " : ""}`,
          }),
        ];
        if (hasTitle) {
          result.push(
            new TextRun({
              text: `${title}`,
            })
          );
        }
        return result;
      }),
    });
  }
  return doc;
}

export function generateIndividualItineraries(
  data: Data,
  messageCollector: MessageCollector
) {
  const { proposedMeetingsData, participantListData, zoomRoomsByName } = data;

  let teamMembers = participantListData.filter(
    (p) => p.teamMemberRoles.length > 0
  );

  let sections = teamMembers.map((teamMember) =>
    generateIndividualItinerarySection(teamMember)
  );

  let doc = new Document({
    sections,
  });

  return doc;

  function generateIndividualItinerarySection(
    individual: Participant
  ): ISectionOptions {
    let meetingsForIndividual = proposedMeetingsData.filter((meeting) => {
      return individual.teamMemberRoles.some((role) => {
        let definition = teamMemberDefinitionsByRole.get(role);
        if (!definition) {
          messageCollector.warn(
            "unable to find definition of team role: " + role
          );
          return false;
        }
        return meeting[definition.property];
      });
    });

    let groupedMeetingsByDate = groupBy(meetingsForIndividual, (m) => m.date);

    return Section({
      headerTitle: "Team Member Itinerary",
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              size: "20pt",
              bold: true,
              text: individual.fullName,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              size: "16pt",
              bold: true,
              text: `${individual.title} Itinerary`,
            }),
          ],
        }),
        new Table({
          layout: "fixed",
          columnWidths: [
            twips.fromInches(1.5),
            twips.fromInches(3),
            twips.fromInches(3),
          ],
          margins: {
            left: twips.fromInches(0.1),
            right: twips.fromInches(0.1),
            top: twips.fromInches(0.01),
            bottom: twips.fromInches(0.01),
            marginUnitType: WidthType.DXA,
          },
          rows: iife(() => {
            let rows: TableRow[] = [];

            for (let group of groupedMeetingsByDate) {
              rows.push(
                new TableRow({
                  children: [
                    HeaderTableCell(group.key, {
                      columnSpan: 3,
                    }),
                  ],
                })
              );

              rows.push(
                new TableRow({
                  children: [
                    HeaderTableCell("Time"),
                    HeaderTableCell("Meeting"),
                    HeaderTableCell("Location"),
                  ],
                })
              );

              for (let meeting of group) {
                rows.push(
                  new TableRow({
                    cantSplit: true,
                    children: [
                      NormalTableCell(meeting.time, {
                        alignment: AlignmentType.RIGHT,
                      }),
                      NormalTableCell(meeting.interviewAssignments),
                      NormalTableCell(
                        iife(() => {
                          let results: (string | IParagraphOptions)[] = [];

                          results.push({
                            // keepNext: !isLast,
                            text: meeting.meetingLocation,
                          });

                          if (meeting.isZoomRoomPrimary) {
                            results.push({
                              // keepNext: !isLast,
                              children: [
                                ZoomLink(
                                  meeting.zoomRoomName,
                                  zoomRoomsByName,
                                  messageCollector
                                ),
                              ],
                            });
                          }

                          return results;
                        })
                      ),
                    ],
                  })
                );
              }
            }
            return rows;
          }),
        }),
      ],
    });
  }
}

function HeaderTableCell(
  text: string,
  options?: { columnSpan?: number; keepNext?: boolean }
) {
  return new TableCell({
    columnSpan: options?.columnSpan,
    shading: {
      fill: "DDDDDD",
      color: "auto",
    },

    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        keepNext: options?.keepNext,
        children: [new TextRun({ bold: true, text })],
      }),
    ],
  });
}

function NormalTableCell(
  children: string | (string | IParagraphOptions)[],
  options?: {
    rowSpan?: number;
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
    hideBottomBorder?: boolean;
    hideTopBorder?: boolean;
    keepNext?: boolean;
  }
) {
  let paragraphOptions: Mutable<IParagraphOptions> = {
    alignment: options?.alignment,
    keepNext: options?.keepNext,
  };

  if (typeof children === "string") {
    children = [children];
  }

  return new TableCell({
    rowSpan: options?.rowSpan,
    borders: {
      bottom: options?.hideBottomBorder
        ? {
            style: BorderStyle.NIL,
          }
        : undefined,
      top: options?.hideTopBorder ? { style: BorderStyle.NIL } : undefined,
    },
    children: children.map((paragraphArgument) => {
      if (typeof paragraphArgument === "string") {
        paragraphArgument = { text: paragraphArgument };
      }

      return new Paragraph({
        ...paragraphOptions,
        ...paragraphArgument,
      });
    }),
  });
}

function separated<T, R>(
  array: T[],
  mapper: (x: T, index: number, array: T[]) => R,
  separator: () => R
) {
  let results: R[] = [];

  array.forEach((item, index) => {
    if (index > 0) {
      results.push(separator());
    }
    results.push(mapper(item, index, array));
  });

  return results;
}

export function generateSummaryItinerary(
  data: Data,
  shouldIncludeRoles: boolean,
  messageCollector: MessageCollector
) {
  let { proposedMeetingsData, zoomRoomsByName } = data;
  let proposedMeetingsGroupedByDate = groupBy(
    proposedMeetingsData,
    (x) => x.date
  );

  let landscape = shouldIncludeRoles;

  let headerTitle: CustomSectionOptions["headerTitle"] = shouldIncludeRoles
    ? "Summary Itinerary with Roles"
    : [
        "Summary Itinerary and Key Contacts",
        "Summary Itinerary & Key Contacts",
      ];

  let columnWidths: number[] = shouldIncludeRoles
    ? [
        twips.fromInches(1.5),
        twips.fromInches(2.5),
        twips.fromInches(2.5),
        twips.fromInches(3.5),
      ]
    : [twips.fromInches(1.5), twips.fromInches(3.0), twips.fromInches(3.0)];

  let doc = new Document({
    sections: [
      Section({
        landscape,
        headerTitle,
        children: [
          ...separated(
            proposedMeetingsGroupedByDate,
            (group) => {
              const proposedMeetingsGroupedByTime = groupBy(
                group,
                (meeting) => meeting.time
              );

              return new Table({
                layout: "fixed",
                columnWidths,
                margins: {
                  left: twips.fromInches(0.1),
                  right: twips.fromInches(0.1),
                  top: twips.fromInches(0.01),
                  bottom: twips.fromInches(0.01),
                  marginUnitType: WidthType.DXA,
                },
                rows: [
                  new TableRow({
                    tableHeader: true,
                    cantSplit: true,
                    children: [
                      HeaderTableCell(group.key, {
                        columnSpan: shouldIncludeRoles ? 4 : 3,
                      }),
                    ],
                  }),
                  new TableRow({
                    tableHeader: true,
                    cantSplit: true,
                    children: [
                      HeaderTableCell("Time"),
                      HeaderTableCell("Meeting"),
                      HeaderTableCell("Location"),
                      ...(shouldIncludeRoles
                        ? [HeaderTableCell("Team Roles")]
                        : []),
                    ],
                  }),
                  ...proposedMeetingsGroupedByTime.flatMap((group) => {
                    const rows: TableRow[] = [];

                    let meetingTime = group.key;

                    group.forEach((meeting, index) => {
                      let row: TableCell[] = [];

                      let isFirst = index === 0;
                      let isLast = index === group.length - 1;

                      row.push(
                        NormalTableCell(isFirst ? meetingTime : "", {
                          alignment: AlignmentType.RIGHT,
                          hideBottomBorder: !isLast,
                          hideTopBorder: !isFirst,
                          keepNext: !isLast,
                        }),
                        NormalTableCell(
                          iife(() => {
                            let results: (string | IParagraphOptions)[] = [];
                            results.push(meeting.interviewAssignments);

                            if (shouldIncludeRoles) {
                              let { teamMembers, isAllTeamMembers } =
                                getParticipantTeamMembers(
                                  data.participantListData,
                                  meeting
                                );

                              let teamMemberStr = isAllTeamMembers
                                ? "All Team Members"
                                : teamMembers.length === 0
                                ? "None"
                                : teamMembers
                                    .map((teamMember) => teamMember.fullName)
                                    .join(", ");

                              results.push({
                                children: [
                                  new TextRun({
                                    italics: true,
                                    size: "8pt",
                                    text: `Team: ${teamMemberStr}`,
                                  }),
                                ],
                              });
                            }

                            return results;
                          })
                        ),
                        NormalTableCell(
                          iife(() => {
                            let results: (string | IParagraphOptions)[] = [];

                            results.push({
                              text: meeting.meetingLocation,
                            });

                            if (meeting.shouldShowZoomRoom) {
                              results.push({
                                children: [
                                  ZoomLink(
                                    meeting.zoomRoomName,
                                    zoomRoomsByName,
                                    messageCollector
                                  ),
                                ],
                              });
                            }

                            return results;
                          })
                        )
                      );
                      if (shouldIncludeRoles) {
                        let generated = false;
                        try {
                          if (meeting.teamRolesData) {
                            let formatted = formatSimpleXml(
                              meeting.teamRolesData
                            );
                            row.push(
                              NormalTableCell([
                                {
                                  children: formatted,
                                },
                              ])
                            );
                            generated = true;
                          }
                        } catch (err) {
                          messageCollector.warn(
                            "err while trying output formatted roles: " +
                              (err instanceof Error ? err.message : err),
                            err
                          );
                        }
                        if (!generated) {
                          row.push(NormalTableCell(meeting.teamRoles));
                          generated = true;
                        }
                      }
                      rows.push(
                        new TableRow({
                          cantSplit: true,
                          children: row,
                        })
                      );
                    });

                    return rows;
                  }),
                ],
              });
            },
            () => new Paragraph("")
          ),
        ],
      }),
    ],
  });

  return doc;
}

type CustomSectionOptions = {
  landscape?: boolean;
  headerTitle: string | [string, string];
  children: ISectionOptions["children"];
};

function Section({
  landscape,
  headerTitle,
  children,
}: CustomSectionOptions): ISectionOptions {
  return {
    properties: {
      page: {
        size: {
          width: landscape ? "11in" : "8.5in",
          height: landscape ? "8.5in" : "11in",
        },
        margin: {
          left: "0.5in",
          right: "0.5in",
          top: "0.5in",
          bottom: "0.5in",
        },
        pageNumbers: {
          start: 1,
        },
      },
      titlePage: true,
    },
    headers: CommonHeader({ title: headerTitle, landscape }),
    children,
  };
}

function CommonHeader({
  mainTitle = "MSCHE Team Visit",
  title,
  landscape = false,
}: {
  mainTitle?: string;
  title: string | [string, string];
  landscape?: boolean;
}): ISectionOptions["headers"] {
  let firstPageHeaderTitle, defaultHeaderTitle;

  if (typeof title === "string") {
    firstPageHeaderTitle = defaultHeaderTitle = title;
  } else {
    [firstPageHeaderTitle, defaultHeaderTitle] = title;
  }

  return {
    first: new Header({
      children: [
        FirstPageHeaderParagraph({
          size: "18pt",
          color: "92002E",
          text: "Commonwealth University",
        }),
        FirstPageHeaderParagraph({
          text: mainTitle,
        }),
        FirstPageHeaderParagraph({
          text: firstPageHeaderTitle,
        }),
        FirstPageHeaderParagraph({
          text: "March 23-26, 2025",
        }),
      ],
    }),
    default: new Header({
      children: [
        DefaultHeaderParagraph({
          landscape,
          children: [
            mainTitle,
            new Tab(),
            { color: "92002E", text: defaultHeaderTitle },
            " | ",
            PageNumber.CURRENT,
          ],
        }),
        DefaultHeaderParagraph({
          landscape,
          text: "",
        }),
      ],
    }),
  };

  function DefaultHeaderParagraph(
    props: HeaderParagraphProps & { landscape: boolean }
  ) {
    let { landscape, ...otherProps } = props;

    return HeaderParagraphCore({
      alignment: AlignmentType.LEFT,
      size: "10pt",
      borders: false,
      spacingAfter: 0,
      tabStops: [
        { type: "left", position: 0 },
        {
          type: "right",
          position: landscape ? twips.fromInches(10) : twips.fromInches(7.5),
        },
      ],

      ...otherProps,
    });
  }

  function FirstPageHeaderParagraph(props: HeaderParagraphProps) {
    return HeaderParagraphCore({
      alignment: AlignmentType.CENTER,
      size: "14pt",
      characterSpacing: twips.fromPoints(1.5),
      borders: true,
      borderColor: "54585A",

      ...props,
    });
  }

  interface HeaderParagraphProps {
    alignment?: IParagraphOptions["alignment"];
    font?: IRunOptions["font"];
    size?: IRunOptions["size"];
    allCaps?: IRunOptions["allCaps"];
    color?: IRunOptions["color"];
    characterSpacing?: IRunOptions["characterSpacing"];
    bold?: IRunOptions["bold"];
    borders?: boolean;
    borderColor?: string;
    spacingAfter?: number;
    tabStops?: IParagraphOptions["tabStops"];

    text?: IRunOptions["text"];
    children?: (
      | string
      | Tab
      | {
          color?: IRunOptions["color"];
          text: IRunOptions["text"];
        }
    )[];
  }
  function HeaderParagraphCore(props: HeaderParagraphProps) {
    let {
      alignment = AlignmentType.CENTER,
      font = "Barlow",
      size,
      allCaps = true,
      color = "54585A",
      characterSpacing,
      bold = true,
      borders,
      borderColor = "54585A",
      spacingAfter = twips.fromPoints(20),
      tabStops,
      text,
      children,
    } = props;

    return new Paragraph({
      alignment,
      spacing: {
        after: spacingAfter,
      },
      tabStops,
      contextualSpacing: true,
      border: borders
        ? {
            top: {
              style: "single",
              size: borderSize.fromPoints(3 / 4),
              color: borderColor,
              space: 8, // points
            },
            bottom: {
              style: "single",
              size: borderSize.fromPoints(3 / 4) / 3,
              color: borderColor,
              space: 8, // points
            },
          }
        : undefined,
      children: children
        ? children.map((child) => {
            let childColor: string | undefined,
              childText: string | Tab | undefined;

            if (typeof child === "string") {
              childColor = undefined;
              childText = child;
            } else if (child instanceof Tab) {
              childColor = undefined;
              childText = child;
            } else {
              ({ color: childColor, text: childText } = child);
            }
            return new TextRun({
              font,
              size,
              allCaps,
              color: childColor ?? color,
              characterSpacing,
              bold,
              children: childText ? [childText] : undefined,
            });
          })
        : [
            new TextRun({
              font,
              size,
              allCaps,
              color,
              characterSpacing,
              bold,
              text,
            }),
          ],
    });
  }
}

function ZoomLink(
  zoomRoomName: string,
  zoomRoomsByName: Map<string, ZoomRoom>,
  messageCollector: MessageCollector
): ParagraphChild {
  let zoomRoom = zoomRoomsByName.get(zoomRoomName);

  if (!zoomRoom) {
    throw new UserError("missing zoom room: " + zoomRoomName);
  }

  let link = zoomRoom.link?.trim() ?? "";
  let hasLink = link !== "";

  if (hasLink) {
    return new ExternalHyperlink({
      link: zoomRoom.link,
      children: [
        new TextRun({
          text: zoomRoom.zoomRoomName,
          style: "Hyperlink",
        }),
      ],
    });
  } else {
    messageCollector.error("zoom room without link: " + zoomRoomName);
    return new TextRun({
      text: zoomRoomName,
    });
  }
}
