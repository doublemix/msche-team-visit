import clsx from "clsx";

export type FormattedHtmlData = Array<{
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
  text: string;
}>;

type Props = {
  data: FormattedHtmlData;
};

export default function FormattedHtml({ data }: Props) {
  return data.map((node, i) => {
    return (
      <span
        key={i}
        className={clsx({
          ["font-bold"]: node.bold,
          ["italic"]: node.italics,
          ["underline"]: node.underline,
        })}
      >
        {node.text}
      </span>
    );
  });
}
