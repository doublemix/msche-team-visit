/** @jsx h
 *  @jsxFrag hFrag */

import xlsx, { type WorkBook } from "xlsx";
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
import * as fs from "fs";

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

function getWorksheets<T extends StructuredRequest<string>>(
  workbook: WorkBook,
  worksheetRequests: T
) {
  return applyStructuredRequest<T, any[], string>(
    worksheetRequests,
    (worksheetName: string) => {
      let worksheet = workbook.Sheets[worksheetName];
      return xlsx.utils.sheet_to_json(worksheet, {
        defval: "",
        skipHidden: true,
      });
    }
  );
}

let boolean = (input: string) => {
  return input.trim() !== "";
};

let stringToBoolean = (yesText: string, noText: string) => (input: string) => {
  if (input === yesText) return true;
  if (input === noText) return false;
  throw new Error(
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
    throw new Error(`Unmapped value for mapInput: ${input}`);
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
    if (!(mapper in data)) throw new Error(`Field ${mapper} not found in data`);
    return data[mapper]?.trim() ?? "";
  } else if (mapper instanceof RegExp) {
    let candidateKeys = Object.keys(data).filter((key) => key.match(mapper));
    if (candidateKeys.length === 0)
      throw new Error(`Field ${mapper} not found in data`);
    if (candidateKeys.length > 1)
      throw new Error(
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
        result[key] = mapField(data, subMapper as any);
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

export type ProposedMeeting = {
  date: string;
  time: string;
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

type Data = {
  proposedMeetingsData: ProposedMeeting[];
  participantListData: Participant[];
  zoomRooms: ZoomRoom[];
  zoomRoomsByName: Map<string, ZoomRoom>;
};

function loadData(filename: string): Data {
  let workbook = xlsx.readFile(filename);

  let [proposedMeetingsRawData, participantListRawData, zoomRoomsRawData] =
    getWorksheets(workbook, [
      "Proposed Meetings-MSCHE Team",
      "Participant List",
      "Zoom Rooms",
    ] as [string, string, string]);
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
    teamMemberRoles: ["Team Member", commaSeparatedList],
  });
  let proposedMeetingsData = mapFields<ProposedMeeting>(
    proposedMeetingsRawData.filter((m) => m.Date),
    {
      date: "Date",
      time: "Time",
      meetingLocation: "Meeting Location",
      // hasZoomRoom: ["Zoom Room Option", stringToBoolean("Yes", "No")],
      zoomRoomOptionType: [
        "Zoom Room Option",
        mapInput([
          ["Primary Room", "primary"],
          ["Yes", "optional"],
          [["No", "N/A", ""], "none"],
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

function generateFullItinerary(data: Data, outputFilename: string) {
  const { proposedMeetingsData, participantListData, zoomRoomsByName } = data;

  let allTeamMembers = participantListData.filter(
    (p) => p.teamMemberRoles.length >= 1
  );

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
                        run: ZoomLink(meeting.zoomRoomName, zoomRoomsByName),
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
                  let teamMembers = participantListData.filter((p) => {
                    for (let d of teamMemberDefinitions) {
                      if (
                        meeting[d.property] &&
                        p.teamMemberRoles.includes(d.value)
                      ) {
                        return true;
                      }
                    }

                    return false;
                  });

                  if (teamMembers.length < 1) {
                    return [];
                  }

                  let isAllTeamMembers =
                    teamMembers.length === allTeamMembers.length;

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
                        console.error(
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
  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputFilename, buffer);
  });
}

function generateIndividualItineraries(data: Data, outputFilename: string) {
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

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputFilename, buffer);
  });

  function generateIndividualItinerarySection(
    individual: Participant
  ): ISectionOptions {
    let meetingsForIndividual = proposedMeetingsData.filter((meeting) => {
      return individual.teamMemberRoles.some((role) => {
        let definition = teamMemberDefinitionsByRole.get(role);
        if (!definition) throw new Error();
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
                                ZoomLink(meeting.zoomRoomName, zoomRoomsByName),
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

function generateSummaryItinerary(data: Data, outputFilename: string) {
  let { proposedMeetingsData, zoomRoomsByName } = data;
  let proposedMeetingsGroupedByDate = groupBy(
    proposedMeetingsData,
    (x) => x.date
  );

  let doc = new Document({
    sections: [
      Section({
        headerTitle: [
          "Summary Itinerary and Key Contacts",
          "Summary Itinerary & Key Contacts",
        ],
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
                columnWidths: [
                  twips.fromInches(1.5),
                  twips.fromInches(3.0),
                  twips.fromInches(3.0),
                ],
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
                    children: [HeaderTableCell(group.key, { columnSpan: 3 })],
                  }),
                  new TableRow({
                    tableHeader: true,
                    cantSplit: true,
                    children: [
                      HeaderTableCell("Time"),
                      HeaderTableCell("Meeting"),
                      HeaderTableCell("Location"),
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
                        NormalTableCell(meeting.interviewAssignments),
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
                                    zoomRoomsByName
                                  ),
                                ],
                              });
                            }

                            return results;
                          })
                        )
                      );
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

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputFilename, buffer);
  });
}

function Section({
  headerTitle,
  children,
}: {
  headerTitle: string | [string, string];
  children: ISectionOptions["children"];
}): ISectionOptions {
  return {
    properties: {
      page: {
        size: {
          width: "8.5in",
          height: "11in",
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
    headers: CommonHeader({ title: headerTitle }),
    children,
  };
}

function CommonHeader({
  mainTitle = "MSCHE Team Visit",
  title,
}: {
  mainTitle?: string;
  title: string | [string, string];
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
          children: [
            mainTitle,
            new Tab(),
            { color: "92002E", text: defaultHeaderTitle },
            " | ",
            PageNumber.CURRENT,
          ],
        }),
        DefaultHeaderParagraph({
          text: "",
        }),
      ],
    }),
  };

  function DefaultHeaderParagraph(props: HeaderParagraphProps) {
    return HeaderParagraphCore({
      alignment: AlignmentType.LEFT,
      size: "10pt",
      borders: false,
      spacingAfter: 0,
      tabStops: [
        { type: "left", position: 0 },
        { type: "right", position: twips.fromInches(7.5) },
      ],

      ...props,
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
  zoomRoomsByName: Map<string, ZoomRoom>
) {
  let zoomRoom = zoomRoomsByName.get(zoomRoomName);

  if (!zoomRoom) {
    throw new Error("missing zoom room: " + zoomRoomName);
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
    console.log("zoom room without link: " + zoomRoomName);
    return new TextRun({
      text: zoomRoomName,
    });
  }
}

function main() {
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

  let data = loadData(filename);

  if (
    !shouldGenerateFullItinerary &&
    !shouldGenerateIndividualItinerary &&
    !shouldGenerateSummaryItinerary
  ) {
    throw new Error(
      "no outputs selected, use --full or --individual or --summary"
    );
  }

  if (shouldGenerateFullItinerary) {
    fullItineraryOutputFile ??= "full-itinerary.docx";
    generateFullItinerary(data, fullItineraryOutputFile);
  }

  if (shouldGenerateIndividualItinerary) {
    individualItineraryOutputFile ??= "individual-itineraries.docx";
    generateIndividualItineraries(data, individualItineraryOutputFile);
  }

  if (shouldGenerateSummaryItinerary) {
    summaryItineraryOutputFile ??= "summary-itinerary.docx";
    generateSummaryItinerary(data, summaryItineraryOutputFile);
  }
}

main();
