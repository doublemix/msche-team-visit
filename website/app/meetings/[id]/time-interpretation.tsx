"use client";

import { useNow } from "@/app/hooks";
import { Clock3 } from "lucide-react";
import { createContext, useCallback, useContext, useState } from "react";

type TimeDisplay = "timeUntil" | "timeDisplay";

let TimeInterpretationContext = createContext<{
  displayType: TimeDisplay;
  toggleDisplayType: () => void;
} | null>(null);

type ProviderProps = {
  defaultDisplay?: TimeDisplay;
  children?: React.ReactNode;
};

export function TimeInterpretationContextProvider({
  defaultDisplay = "timeUntil",
  children,
}: ProviderProps) {
  let [displayType, setDisplayType] = useState(defaultDisplay);

  let toggleDisplayType = useCallback(() => {
    setDisplayType((previous) =>
      previous === "timeDisplay" ? "timeUntil" : "timeDisplay",
    );
  }, []);

  return (
    <TimeInterpretationContext value={{ displayType, toggleDisplayType }}>
      {children}
    </TimeInterpretationContext>
  );
}

type TimeInterpretationProps = {
  startTime: Date | null;
  endTime: Date | null;
  timeDisplay: React.ReactNode;
  defaultDisplay?: TimeDisplay;
};

export default function TimeInterpretation({
  defaultDisplay,
  ...otherProps
}: TimeInterpretationProps) {
  let contextValue = useContext(TimeInterpretationContext);

  if (contextValue === null) {
    return (
      <TimeInterpretationContextProvider defaultDisplay={defaultDisplay}>
        <TimeInterpretation {...otherProps} />
      </TimeInterpretationContextProvider>
    );
  }

  let { displayType, toggleDisplayType } = contextValue;

  return (
    <TimeInterpretationCore
      {...otherProps}
      displayType={displayType}
      onToggleDisplayType={toggleDisplayType}
    />
  );
}

type TimeInterpretationCoreProps = {
  startTime: Date | null;
  endTime: Date | null;
  timeDisplay: React.ReactNode;
  displayType: TimeDisplay;
  onToggleDisplayType: () => void;
};
function TimeInterpretationCore({
  startTime,
  endTime,
  timeDisplay,
  displayType,
  onToggleDisplayType,
}: TimeInterpretationCoreProps) {
  let now = useNow();
  let message = getTimeInterpretation(now, startTime, endTime);

  if (message === null) return null;

  return (
    <div onClick={onToggleDisplayType}>
      {displayType === "timeUntil" && (
        <>
          <Clock3 size="0.9em" className="inline align-baseline" /> {message}
        </>
      )}
      {displayType === "timeDisplay" && <>{timeDisplay}</>}
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
    let afterStart = now - startTimeTs;
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
