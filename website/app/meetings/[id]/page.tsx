import { singleOrDefault } from "@/app/utils";
import { getDatabase } from "@/database";
import { notFound } from "next/navigation";
import TimeInterpretation from "./time-interpretation";
import FormattedHtml, { FormattedHtmlData } from "@/app/formatted-html";

type Meeting = {
  id: number;
  date: string;
  timeDisplay: string;
  startTime: Date | null;
  endTime: Date | null;
  name: string;
  location: string;
  zoomRoomId: number | null;
  roleInfo: FormattedHtmlData | null;
};

type Participant = {
  id: number;
  name: string;
  title: string;
  isMsche: boolean;
};

type ZoomMeetingRoom = {
  id: number;
  name: string;
  link: string;
};

async function getData(id: string) {
  let idAsNumber = parseInt(id);

  if (!isFinite(idAsNumber)) notFound();

  let sql = getDatabase();

  let meetingData = await sql`
      SELECT id,
        date, time as "timeDisplay", starttime as "startTime", endtime as "endTime",
        name, location, zoomRoomId as "zoomRoomId", "roleInfo"
      FROM meetings WHERE id = ${idAsNumber}`;

  let meeting = singleOrDefault(meetingData, null) as Meeting;

  if (meeting === null) notFound();

  let participants = (await sql`
    SELECT p.id, p.name, p.title, p."isMsche"
    FROM meetingParticipation mp
    JOIN participants p ON mp.participantId = p.id
    WHERE mp.meetingId = ${idAsNumber}`) as Participant[];

  let zoomRoom: ZoomMeetingRoom | null = null;
  if (meeting.zoomRoomId !== null) {
    let zoomRoomData =
      await sql`SELECT name, link FROM zoomRooms WHERE id = ${meeting.zoomRoomId}`;

    zoomRoom = singleOrDefault(zoomRoomData as ZoomMeetingRoom[], null);
  }

  let representatives = participants.filter((p) => !p.isMsche);
  let teamMembers = participants.filter((p) => p.isMsche);

  return {
    meeting,
    representatives,
    teamMembers,
    zoomRoom,
  };
}

type Props = {
  params: Promise<{ id: string }>;
};

export default async function Page({ params }: Props) {
  let { id } = await params;

  let { meeting, representatives, teamMembers, zoomRoom } = await getData(id);

  return (
    <div className="flex flex-col gap-y-2">
      <h1 className="text-3xl font-bold">{meeting.name}</h1>
      <div className="flex flex-col flex-nowrap gap-0 text-sm/[120%] text-gray-500">
        <div>
          {meeting.date}, {meeting.timeDisplay}
        </div>
        <div>{meeting.location}</div>
        <TimeInterpretation
          startTime={meeting.startTime}
          endTime={meeting.endTime}
          timeDisplay={meeting.timeDisplay}
        />
      </div>
      {zoomRoom && (
        <div>
          <span className="font-bold">Zoom Room Option:</span>{" "}
          <a className="text-blue-500" href={zoomRoom.link}>
            {zoomRoom.name}
          </a>
        </div>
      )}
      <ParticipantList title="MSCHE Team Members" participants={teamMembers} />
      <ParticipantList
        title="CU Representatives"
        participants={representatives}
      />
      {meeting.roleInfo && (
        <div>
          <div className="leading-[110%] font-bold">Role Info:</div>
          <div className="pl-4 text-sm leading-[110%] text-gray-600">
            <FormattedHtml data={meeting.roleInfo} />
          </div>
        </div>
      )}
    </div>
  );
}

type ParticipantListProps = {
  title: React.ReactNode;
  participants: Participant[];
};
function ParticipantList({ title, participants }: ParticipantListProps) {
  return (
    <div>
      <div className="font-bold">{title}:</div>
      <ul className="list-disc pl-8">
        {participants.length === 0 && <li className="italic">None</li>}
        {participants.map((p) => {
          return (
            <li key={p.id}>
              {p.name}
              <span className="inline-block w-4"> </span>
              <span className="text-xs text-gray-500">{p.title}</span>
            </li>
          );
        })}
      </ul>
    </div>
  );
}
