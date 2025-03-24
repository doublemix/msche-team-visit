"use client";

import { useNow } from "@/app/hooks";
import { Clock3 } from "lucide-react";

type TimeInterpretationProps = {
  startTime: Date | null;
  endTime: Date | null;
};
export default function TimeInterpretation({
  startTime,
  endTime,
}: TimeInterpretationProps) {
  let now = useNow();
  let message = getTimeInterpretation(now, startTime, endTime);
  if (message === null) return null;

  return (
    <div>
      <Clock3 size={14} className="inline align-top" /> {message}
    </div>
  );
}

const SECOND_DURATION = 1000;
const MINUTE_DURATION = 60 * SECOND_DURATION;
const HOUR_DURATION = 60 * MINUTE_DURATION;
const DAY_DURATION = 24 * HOUR_DURATION;
function getTimeInterpretation(
  now: number,
  startTime: Date | null,
  endTime: Date | null,
) {
  let startTimeTs = startTime?.valueOf() ?? null;
  let endTimeTs = endTime?.valueOf() ?? null;

  if (startTimeTs !== null && now < startTimeTs) {
    let beforeStart = startTimeTs - now;
    let durationDisplay = formatDurationMs(beforeStart);

    return `Starts in ${durationDisplay}`;
  }
  if (endTimeTs !== null && now > endTimeTs) {
    let afterEnd = now - endTimeTs;
    let durationDisplay = formatDurationMs(afterEnd);

    return `Ended ${durationDisplay} ago`;
  }

  if (
    startTimeTs !== null &&
    endTimeTs !== null &&
    startTimeTs <= now &&
    endTimeTs >= now
  ) {
    return "Ongoing";
  }

  if (startTimeTs !== null && endTimeTs === null) {
    let afterStart = startTimeTs - now;
    let durationDisplay = formatDurationMs(afterStart);

    return `${durationDisplay} ago`;
  }

  if (startTimeTs === null && endTimeTs !== null) {
    let beforeEnd = endTimeTs - now;
    let durationDisplay = formatDurationMs(beforeEnd);

    return `Ends in ${durationDisplay}`;
  }

  return null;
}
function formatDurationMs(duration: number) {
  let days = Math.floor(duration / DAY_DURATION);
  let hours = Math.floor(duration / HOUR_DURATION);
  let minutes = Math.floor(duration / MINUTE_DURATION);
  let seconds = Math.floor(duration / SECOND_DURATION);

  let durationDisplay: string;
  if (days >= 1) durationDisplay = format(days, "day", "days");
  else if (hours >= 1) durationDisplay = format(hours, "hour", "hours");
  else if (minutes >= 1) durationDisplay = format(minutes, "minute", "minutes");
  else durationDisplay = format(seconds, "second", "seconds");

  return durationDisplay;
}
function format(count: number, singular: string, plural: string) {
  return `${count} ${count === 1 ? singular : plural}`;
}
