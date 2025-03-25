import Link from "next/link";
import { getDatabase } from "@/database";
import { groupBy } from "../utils";
import { Fragment } from "react";
import TimeInterpretation, {
  TimeInterpretationContextProvider,
} from "./[id]/time-interpretation";

type Meeting = {
  id: number;
  name: string;
  dateDisplay: string;
  timeDisplay: string;
  startTime: Date | null;
  endTime: Date | null;
};

async function getData() {
  let sql = getDatabase();

  let data = await sql(
    `SELECT id, name, date as "dateDisplay", time as "timeDisplay", startTime as "startTime", endTime as "endTime" FROM meetings`,
  );

  return data as Meeting[];
}

export default async function Page() {
  let meetings = await getData();

  let meetingsGroupedByDate = groupBy(
    meetings,
    (meeting) => meeting.dateDisplay,
  );

  return (
    <TimeInterpretationContextProvider defaultDisplay="timeDisplay">
      <div className="overflow-auto">
        <table className="table-auto">
          <tbody>
            {meetingsGroupedByDate.map((meetingGroup) => {
              return (
                <Fragment key={meetingGroup.key}>
                  <tr>
                    <td
                      colSpan={2}
                      className="bg-gray-300 px-4 py-2 text-lg font-bold"
                    >
                      {meetingGroup.key}
                    </td>
                  </tr>
                  {meetingGroup.map((meeting) => {
                    return (
                      <tr
                        key={meeting.id}
                        className="odd:bg-white even:bg-gray-100"
                      >
                        <td className="px-4 py-2">
                          <Link
                            className="text-blue-400"
                            href={`/meetings/${meeting.id}`}
                          >
                            {meeting.name}
                          </Link>
                        </td>
                        <td className="px-4 py-2 text-right whitespace-nowrap">
                          <TimeInterpretation
                            defaultDisplay="timeDisplay"
                            timeDisplay={meeting.timeDisplay}
                            startTime={meeting.startTime}
                            endTime={meeting.endTime}
                          />
                        </td>
                      </tr>
                    );
                  })}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </TimeInterpretationContextProvider>
  );
}
