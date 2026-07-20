import * as React from "react";
import {
  IPersonaSharedProps,
  Persona,
  PersonaSize,
} from "@fluentui/react/lib/Persona";
import { Text } from "@fluentui/react/lib/Text";
import { IPersonProps } from "./IPersonProps";
import { getUserPhotoUrl } from "../../services/PhotoService";

export const Person: React.FunctionComponent<IPersonProps> = (
  props: IPersonProps
) => {
  const { text, secondaryText, userEmail, size, tertiaryText, graphClient } =
    props;

  // Start with the classic SharePoint profile-photo endpoint as an immediate
  // fallback, then swap in the Entra ID (Microsoft Graph) photo once/if it
  // resolves — that's the actual source of truth for photos in this tenant.
  const [imageUrl, setImageUrl] = React.useState<string | undefined>(
    userEmail
      ? `/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(userEmail)}`
      : undefined
  );

  React.useEffect(() => {
    let cancelled = false;

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    getUserPhotoUrl(graphClient, userEmail).then((graphPhotoUrl) => {
      if (!cancelled && graphPhotoUrl) {
        setImageUrl(graphPhotoUrl);
      }
    });

    return () => {
      cancelled = true;
    };
  }, [graphClient, userEmail]);

  const personProps: IPersonaSharedProps = React.useMemo(() => {
    return {
      imageUrl,
      text: text,
      secondaryText: secondaryText,
      tertiaryText: tertiaryText,
    };
  }, [imageUrl, text, secondaryText, tertiaryText]);

  const _onRenderPrimaryText = React.useCallback(() => {
    return (
      <>
        <Text
          title={text}
          variant="mediumPlus"
          block
          nowrap
          styles={{ root: { fontWeight: 600 } }}
        >
          {text}
        </Text>
      </>
    );
  }, [text]);

  const _onRenderSecondaryText = React.useCallback(() => {
    return (
      <>
        <Text
          title={secondaryText}
          variant="smallPlus"
          block
          nowrap
          styles={{ root: { fontWeight: 400 } }}
        >
          {secondaryText}
        </Text>
      </>
    );
  }, [secondaryText]);



  return (
    <>
      <Persona
        {...personProps}
        size={size || PersonaSize.size40}
        onRenderPrimaryText={_onRenderPrimaryText}
        onRenderSecondaryText={_onRenderSecondaryText}
        styles={{
          secondaryText: { maxWidth: 230 },
          primaryText: { maxWidth: 230 },
        }}
      />
    </>
  );
};