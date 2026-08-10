import { mergeStyleSets } from "@fluentui/react";

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
const gridClasses = mergeStyleSets({
  grid: {
    display: "flex",
    flexWrap: "wrap",
    gap: 16,
  },
  groupHeading: {
    fontWeight: 600,
  },
});

// eslint-disable-next-line @typescript-eslint/explicit-module-boundary-types, @typescript-eslint/explicit-function-return-type
export const usePositionDirectoryStyles = () => {
  return { gridClasses };
};
