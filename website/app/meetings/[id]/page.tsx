import { singleOrDefault } from "@/app/utils";
import { getDatabase } from "@/database";
import { notFound } from "next/navigation";
import TimeInterpretation from "./time-interpretation";

type Meeting = {
  id: number;
  date: string;
  time: string;
  startTime: Date | null;
  endTime: Date | null;
  name: string;
  location: string;
  zoomRoomId: number | null;
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

  let meetingData =
    await sql`SELECT id, date, time, starttime as "startTime", endtime as "endTime", name, location, zoomRoomId as "zoomRoomId" FROM meetings WHERE id = ${idAsNumber}`;

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
          {meeting.date}, {meeting.time}
        </div>
        <div>{meeting.location}</div>
        <TimeInterpretation
          startTime={meeting.startTime}
          endTime={meeting.endTime}
        />
      </div>
      {zoomRoom && (
        <div>
          Zoom Room Option:{" "}
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
      <div>{title}:</div>
      <ul className="list-disc pl-8">
        {participants.map((p) => {
          return (
            <li key={p.id}>
              {p.name}{" "}
              <span className="inline-block pl-1 text-xs text-gray-500">
                {p.title}
              </span>
            </li>
          );
        })}
      </ul>
    </div>
  );
}
