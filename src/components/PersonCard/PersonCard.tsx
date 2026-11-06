/* tslint:disable */
import * as React from "react";
import {
  DocumentCard,
  DocumentCardActions,
  DocumentCardDetails,
  IDocumentCard,
  IDocumentCardActionsStyles,
} from "@fluentui/react/lib/DocumentCard";

import {
  Stack,
  IButtonProps,
  IButtonStyles,
  PersonaSize,
} from "@fluentui/react";
import { useOrgChartColors } from "../../common/OrgChartThemeContext";
import { LivePersona } from "@pnp/spfx-controls-react/lib/controls/LivePersona";
import { Person } from "../Person/Person";
import { usePersonaCardStyles } from "./usePersonaCardStyles";
import { IPersonCardProps } from "./IPersonCardProps";

export const PersonCard: React.FunctionComponent<IPersonCardProps> = (
  props: IPersonCardProps
) => {
  const { userInfo, onUserSelected, showActionsBar, graphClient, serviceScope } = props;

  const documentCardRef = React.useRef<IDocumentCard>(null);
  const {
    personaCardStyles,
    documentCardActionStyles: baseActionStyles,
    buttonStyles: baseButtonStyles,
    stackPersonaStyles,
  } = usePersonaCardStyles();

  // Admin-configured colours (undefined = keep the SharePoint theme look).
  const { cardBackgroundColor, textColor } = useOrgChartColors();

  const documentCardActionStyles: Partial<IDocumentCardActionsStyles> =
    React.useMemo(
      () =>
        cardBackgroundColor
          ? {
              ...baseActionStyles,
              root: {
                ...(baseActionStyles.root as object),
                backgroundColor: cardBackgroundColor,
              },
            }
          : baseActionStyles,
      [baseActionStyles, cardBackgroundColor]
    );

  const buttonStyles: IButtonStyles = React.useMemo(
    () =>
      textColor
        ? {
            ...baseButtonStyles,
            icon: { ...(baseButtonStyles.icon as object), color: textColor },
            iconHovered: {
              ...(baseButtonStyles.iconHovered as object),
              color: textColor,
            },
          }
        : baseButtonStyles,
    [baseButtonStyles, textColor]
  );

  const documentCardActions: IButtonProps[] = React.useMemo(() => {
    const cardActions: IButtonProps[] = [];

    cardActions.push({
      iconProps: { iconName: "Chat" },
      title: "Chat",
      styles: buttonStyles,
      onClick: (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        window.open(
          `https://teams.microsoft.com/l/chat/0/0?users=${userInfo.email}&message=Hi ${userInfo.displayName} `,
          "_blank"
        );
      },
    });

    if (userInfo?.email) {
      cardActions.push({
        iconProps: { iconName: "Mail" },
        title: "Mail",
        styles: buttonStyles,
        onClick: (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          window.open(`MAILTO:${userInfo.email}`, "_blank");
        },
      });
    }

    if (userInfo?.workPhone) {
      cardActions.push({
        iconProps: { iconName: "Phone" },
        title: "Call",
        styles: buttonStyles,
        onClick: (ev) => {
          ev.preventDefault();
          ev.stopPropagation();
          window.open(`CALLTO:${userInfo.workPhone}`, "_blank");
        },
      });
    }

    if (userInfo.hasDirectReports) {
      cardActions.push({
        iconProps: { iconName: "Org" },
        title: "View Organization",
        styles: { ...buttonStyles },
      });
    }
    return cardActions;
  }, [
    buttonStyles,
    userInfo.displayName,
    userInfo.email,
    userInfo.hasDirectReports,
    userInfo.workPhone,
  ]);

  return (
    <>
      <DocumentCard
        className={personaCardStyles.tile}
        componentRef={documentCardRef}
        // Inline style so the admin colour always wins over the class-based
        // default background from usePersonaCardStyles.
        style={
          cardBackgroundColor || textColor
            ? {
                ...(cardBackgroundColor
                  ? { backgroundColor: cardBackgroundColor }
                  : {}),
                ...(textColor ? { color: textColor } : {}),
              }
            : undefined
        }
        onClick={() => {
          if (userInfo.hasDirectReports) {
            onUserSelected(userInfo);
          }
        }}
      >
        <DocumentCardDetails>
          <Stack
            horizontal
            horizontalAlign="space-between"
            styles={stackPersonaStyles}
          >
            {/* LivePersona wraps our own card content with Microsoft's
                native persona card — presence, availability, working
                hours, local time, contact, and "reports to" — the same
                one Teams/Outlook/the native org chart use. No custom
                Graph calls needed for any of that; it uses SharePoint's
                built-in personCard experience via serviceScope. */}
            <LivePersona
              upn={userInfo.email}
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              serviceScope={serviceScope as any}
              template={
                <Person
                  text={userInfo.displayName}
                  secondaryText={userInfo.title}
                  userEmail={userInfo.email}
                  pictureUrl={userInfo.pictureUrl}
                  graphClient={graphClient}
                  size={PersonaSize.size48}
                />
              }
            />
          </Stack>
        </DocumentCardDetails>
        {showActionsBar && (
          <DocumentCardActions
            actions={documentCardActions}
            styles={documentCardActionStyles}
          />
        )}
      </DocumentCard>
    </>
  );
};