import * as React from "react";
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import { groupBy, sortBy } from "lodash";
import { PersonCard } from "../PersonCard/PersonCard";
import { IUserInfo } from "../../models/IUserInfo";
import { getUsersByJobTitle } from "../../services/PeopleSearchService";
import { IPositionDirectoryProps } from "./IPositionDirectoryProps";
import { usePositionDirectoryStyles } from "./usePositionDirectoryStyles";

// Clicking a PersonCard in the org chart drills into that person's own
// direct reports. There's no tree here to drill into, so clicks are a
// deliberate no-op — the card just displays.
const noopSelect = (): void => undefined;

export const PositionDirectory: React.FunctionComponent<
  IPositionDirectoryProps
> = (props: IPositionDirectoryProps) => {
  const { context, jobTitles, groupByDepartment, showActionsBar, graphClient, sp } =
    props;
  const { gridClasses } = usePositionDirectoryStyles();

  const [people, setPeople] = React.useState<IUserInfo[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [hasError, setHasError] = React.useState<boolean>(false);

  const jobTitlesKey: string = (jobTitles ?? []).join("|");

  React.useEffect(() => {
    let cancelled = false;
    setIsLoading(true);
    setHasError(false);

    getUsersByJobTitle(sp, jobTitles ?? [])
      .then((result) => {
        if (!cancelled) setPeople(result);
      })
      .catch(() => {
        if (!cancelled) setHasError(true);
      })
      .finally(() => {
        if (!cancelled) setIsLoading(false);
      });

    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sp, jobTitlesKey]);

  if (!jobTitles || jobTitles.length === 0) {
    return (
      <Stack horizontalAlign="center" styles={{ root: { padding: 20 } }}>
        <Text>Bitte in den Web Part-Eigenschaften eine Position angeben.</Text>
      </Stack>
    );
  }

  if (isLoading) {
    return (
      <Stack
        style={{ minHeight: 200 }}
        verticalAlign="center"
        horizontalAlign="center"
      >
        <Spinner
          size={SpinnerSize.large}
          label="Lade Übersicht..."
          labelPosition="bottom"
        />
      </Stack>
    );
  }

  if (hasError) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        Personen konnten nicht geladen werden.
      </MessageBar>
    );
  }

  if (people.length === 0) {
    return (
      <Stack horizontalAlign="center" styles={{ root: { padding: 20 } }}>
        <Text>Keine Personen mit dieser Position gefunden.</Text>
      </Stack>
    );
  }

  const renderCard = (person: IUserInfo): JSX.Element => (
    <PersonCard
      key={person.id ?? person.email}
      userInfo={person}
      onUserSelected={noopSelect}
      selectedUser={undefined}
      showActionsBar={showActionsBar}
      graphClient={graphClient}
      serviceScope={context.serviceScope}
      sp={sp}
    />
  );

  if (!groupByDepartment) {
    return (
      <Stack styles={{ root: { padding: 20 } }}>
        <div className={gridClasses.grid}>{people.map(renderCard)}</div>
      </Stack>
    );
  }

  const groups = groupBy(people, (p) => p.department || "Ohne Abteilung");
  const departmentNames = Object.keys(groups).sort((a, b) =>
    a.localeCompare(b)
  );

  return (
    <Stack tokens={{ childrenGap: 24 }} styles={{ root: { padding: 20 } }}>
      {departmentNames.map((department) => (
        <Stack key={department} tokens={{ childrenGap: 8 }}>
          <Text variant="large" className={gridClasses.groupHeading}>
            {department}
          </Text>
          <div className={gridClasses.grid}>
            {sortBy(groups[department], "displayName").map(renderCard)}
          </div>
        </Stack>
      ))}
    </Stack>
  );
};
