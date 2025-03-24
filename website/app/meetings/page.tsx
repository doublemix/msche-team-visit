import Link from "next/link";
import { getDatabase } from "@/database";

type Meeting = {
  id: number;
  name: string;
};

async function getData() {
  let sql = getDatabase();

  let data = await sql(`SELECT id, name FROM meetings`);

  return data as Meeting[];
}

export default async function Page() {
  let meetings = await getData();

  return (
    <>
      <table className="table-auto">
        <tbody>
          {meetings.map((meeting) => {
            return (
              <tr key={meeting.id} className="odd:bg-white even:bg-gray-100">
                <td className="px-4 py-2">
                  <Link
                    className="text-blue-400"
                    href={`/meetings/${meeting.id}`}
                  >
                    {meeting.name}
                  </Link>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </>
  );
}
