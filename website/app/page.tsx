import { getDatabase } from "@/database";
import { getNowFromSearch } from "./utils";
import React from "react";
import TimeInterpretation, {
  TimeInterpretationContextProvider,
} from "./meetings/[id]/time-interpretation";
import Link from "next/link";

interface Meeting {
  id: number;
  name: string;
  timeDisplay: string;
  startTime: Date | null;
  endTime: Date | null;
}
async function getData(now: number) {
  let sql = getDatabase();

  let date = new Date(now);

  let currentMeetings = (await sql`
    SELECT
      id, name, time as "timeDisplay", startTime as "startTime", endTime as "endTime"
    FROM meetings
    WHERE startTime IS NOT NULL
      AND endTime IS NOT NULL
      AND startTime <= ${date}
      AND ${date} <= endTime
  `) as Meeting[];

  let upcomingMeetings = (await sql`
    SELECT
      id, name, time as "timeDisplay", startTime as "startTime", endTime as "endTime"
    FROM meetings
    WHERE startTime IS NOT NULL
      AND startTime = (SELECT MIN(startTime) FROM meetings WHERE startTime > ${date})
  `) as Meeting[];

  return {
    currentMeetings,
    upcomingMeetings,
  };
}

interface PageProps {
  searchParams: Promise<{ now?: string }>;
}
export default async function Page({ searchParams }: PageProps) {
  let now = getNowFromSearch(await searchParams);
  let { currentMeetings, upcomingMeetings } = await getData(now);

  return (
    <div className="flex flex-col gap-8">
      <TimeInterpretationContextProvider defaultDisplay="timeDisplay">
        <MeetingList title="Current Meetings" meetings={currentMeetings} />
      </TimeInterpretationContextProvider>
      <TimeInterpretationContextProvider defaultDisplay="timeUntil">
        <MeetingList title="Upcoming Meetings" meetings={upcomingMeetings} />
      </TimeInterpretationContextProvider>
    </div>
  );
}

type MeetingListProps = {
  title: React.ReactNode;
  meetings: Meeting[];
};
function MeetingList({ title, meetings }: MeetingListProps) {
  return (
    <div className="flex flex-col gap-2">
      <div className="text-sm tracking-widest text-gray-500 uppercase underline decoration-gray-200">
        {title}
      </div>
      {meetings.length > 0 ? (
        <div className="flex flex-col gap-4">
          {meetings.map((m) => {
            return (
              <div
                className="flex flex-col gap-1 rounded-xl border border-gray-200 p-4 leading-[100%] shadow"
                key={m.id}
              >
                <div>
                  <Link className="text-blue-500" href={`/meetings/${m.id}`}>
                    {m.name}
                  </Link>
                </div>
                <div className="text-sm/[120%] text-gray-500">
                  <TimeInterpretation
                    startTime={m.startTime}
                    endTime={m.endTime}
                    timeDisplay={m.timeDisplay}
                  />
                </div>
              </div>
            );
          })}
        </div>
      ) : (
        <div className="italic">None</div>
      )}
    </div>
  );
}
