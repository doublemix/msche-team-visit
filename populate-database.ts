import { neon } from "@neondatabase/serverless";
import {
  loadData,
  type ProposedMeeting,
  type Data,
  type TimeOfDay,
  teamMemberDefinitions,
  MessageCollector,
  formatSimpleXml,
  type FormattingOptions,
} from "./generate-documents.ts";
import fs from "node:fs";
import { _throw, single } from "./website/app/utils.ts";

function formatTimestamp(date: Date | null, time: TimeOfDay | null) {
  if (date === null) return null;
  if (time === null) return null;
  return `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()} ${
    time.hour24
  }:${time.minute}-04:00`;
}

async function populateDatabase(data: Data) {
  let sql = neon(`${process.env.DATABASE_URL}`);

  await sql(`DROP TABLE IF EXISTS meetingParticipation`);
  await sql(`DROP TABLE IF EXISTS participants`);
  await sql(`DROP TABLE IF EXISTS meetings`);
  await sql(`DROP TABLE IF EXISTS zoomRooms`);

  await sql(`CREATE TABLE zoomRooms(
    id INTEGER PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
    name text,
    link text
  )`);

  await sql(`CREATE TABLE meetings(
    id INTEGER PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
    date text,
    time text,
    startTime timestamp with time zone,
    endTime timestamp with time zone,
    name text,
    location text,
    zoomRoomId integer null references zoomRooms(id),
    "roleInfo" jsonb
  )`);

  await sql(`CREATE TABLE participants(
    id INTEGER PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
    name text,
    title text,
    "isMsche" bool not null
  )`);

  await sql(`CREATE TABLE meetingParticipation(
    id INTEGER PRIMARY KEY GENERATED ALWAYS AS IDENTITY,
    meetingId integer NOT NULL REFERENCES meetings(id),
    participantId integer NOT NULL REFERENCES participants(id)
  )`);

  let zoomRoomIds = new Map<string, number>();

  for (let zoomRoom of data.zoomRooms) {
    let result = await sql`INSERT INTO zoomRooms(
      name,
      link
      ) VALUES (
       ${zoomRoom.zoomRoomName},
       ${zoomRoom.link}
      ) RETURNING id`;
    let zoomRoomId = single(result).id;
    zoomRoomIds.set(zoomRoom.zoomRoomName, zoomRoomId);
  }

  let meetingIds = new Map<ProposedMeeting, number>();
  for (let meeting of data.proposedMeetingsData) {
    let zoomRoomId: number | null = null;
    let zoomRoomName = meeting.zoomRoomName;

    if (zoomRoomName) {
      zoomRoomId =
        zoomRoomIds.get(zoomRoomName) ?? _throw(new Error("unknown zoom room"));
    }

    let roleInfo =
      meeting.teamRolesData !== null
        ? formatSimpleXml(
            meeting.teamRolesData,
            (text: string, formattingOptions: FormattingOptions) => ({
              ...formattingOptions,
              text,
            })
          )
        : null;

    let result = await sql`
      INSERT INTO meetings(
        date,
        time,
        startTime,
        endTime,
        name,
        location,
        zoomRoomId,
        "roleInfo"
      ) VALUES (
        ${meeting.date},
        ${meeting.time},
        ${formatTimestamp(meeting.parsedDate, meeting.startTime)},
        ${formatTimestamp(meeting.parsedDate, meeting.endTime)},
        ${meeting.interviewAssignments},
        ${meeting.meetingLocation},
        ${zoomRoomId},
        ${JSON.stringify(roleInfo)}
      ) RETURNING id`;

    let meetingId = single(result).id;

    meetingIds.set(meeting, meetingId);
  }

  let participantIds = new Map<string, number>();
  for (let participant of data.participantListData) {
    let result = await sql`
      INSERT INTO participants(
        name,
        title,
        "isMsche"
      ) VALUES (
       ${participant.fullName},
       ${participant.title},
       ${participant.teamMemberRoles.length > 0}
      ) RETURNING id;`;

    let participantId = single(result).id;

    participantIds.set(participant.id, participantId);
  }

  let mscheTeamMembers = data.participantListData.filter(
    (x) => x.teamMemberRoles.length > 0
  );

  for (let meeting of data.proposedMeetingsData) {
    let meetingId =
      meetingIds.get(meeting) ?? _throw(new Error("missing meeting"));

    for (let individual of meeting.individuals) {
      let participantId = participantIds.get(individual.id);
      // ??_throw(new Error("missing individual: " + individual.displayName));

      if ((participantId ?? null) === null) continue;

      await sql`
        INSERT INTO meetingParticipation(meetingId, participantId)
        VALUES (${meetingId}, ${participantId})
      `;
    }

    for (let mscheTeamMember of mscheTeamMembers) {
      let participantId = participantIds.get(mscheTeamMember.id);
      if (
        mscheTeamMember.teamMemberRoles.some((role) => {
          let teamMemberDefinition = teamMemberDefinitions.find(
            (x) => x.value === role
          );
          if (!teamMemberDefinition) return false;
          return meeting[teamMemberDefinition.property];
        })
      ) {
        await sql`
          INSERT INTO meetingParticipation(meetingId, participantId)
          VALUES (${meetingId}, ${participantId})
        `;
      }
    }
  }

  console.log("Successfully populated database");
}

let sourceFile = fs.readFileSync(process.argv[2]);
let data = loadData(
  sourceFile,
  {
    teamRoleSource: { type: "meetingsTable", nameRow: 0, headerRow: 2 },
    meetingRange: 2,
  },
  new MessageCollector()
);
populateDatabase(data);
