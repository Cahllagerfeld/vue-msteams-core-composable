import * as microsoftTeams from "@microsoft/teams-js";
import { Ref, ref, watchEffect } from "vue";

export interface useTeamsResult {
  isInTeams: Ref<boolean>;
  context: Ref<microsoftTeams.Context | undefined>;
}

export function checkTeams(): boolean {
  const microsoftTeamsLib = microsoftTeams || window["microsoftTeams"];

  if (!microsoftTeamsLib) {
    return false;
  }

  if (
    (window.parent === window.self && (window as any).nativeInterface) ||
    window.name === "embedded-page-container" ||
    window.name === "extension-tab-frame"
  ) {
    return true;
  }
  return false;
}

export function useTeams(): useTeamsResult {
  const context = ref<microsoftTeams.Context | undefined>(undefined);
  const isInTeams = ref<boolean>(false);
  watchEffect(() => {
    isInTeams.value = checkTeams();
    if (isInTeams.value === true) {
      microsoftTeams.initialize(() => {
        microsoftTeams.getContext((msTeamsContext: microsoftTeams.Context) => {
          context.value = msTeamsContext;
        });
      });
    } else {
      microsoftTeams.initialize();
    }
  });
  return { isInTeams, context };
}
