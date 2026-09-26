import { createFileRoute } from "@tanstack/react-router";
import { useState } from "react";

import { useCustomThemes } from "../hooks/useCustomThemes";
import { useTheme } from "../hooks/useTheme";
import { ThemeLibrary } from "../components/settings/ThemeSettings";
import { UsageLimitsSection } from "../components/usage/UsageLimits";
import { WorkspacePageContainer } from "../components/WorkspacePageContainer";
import { WorkspacePageHeader } from "../components/WorkspacePageHeader";

function AccountRoute() {
  const {
    appearanceMode,
    refreshTheme,
    resolvedTheme,
    setAppearanceMode,
    setTheme,
    setThemeHalf,
    theme,
    themeHalves,
  } = useTheme();
  const customThemes = useCustomThemes();
  const [isImportOpen, setIsImportOpen] = useState(false);
  const limitsNow = Date.now();

  return (
    <>
      <WorkspacePageHeader>
        <h1 className="text-sm font-medium">Account</h1>
      </WorkspacePageHeader>
      <div className="topbar-scroll-fade scrollbar-gutter-both min-h-0 flex-1 overflow-y-auto">
        <WorkspacePageContainer className="gap-8">
          <section aria-labelledby="account-themes-heading" className="flex flex-col gap-4">
            <h1 id="account-themes-heading" className="text-lg font-semibold">
              Themes
            </h1>
            <ThemeLibrary
              theme={theme}
              setTheme={setTheme}
              appearanceMode={appearanceMode}
              setAppearanceMode={setAppearanceMode}
              customThemes={customThemes}
              initialAppearance={resolvedTheme}
              refreshTheme={refreshTheme}
              isImportOpen={isImportOpen}
              onImportOpenChange={setIsImportOpen}
              themeHalves={themeHalves}
              setThemeHalf={setThemeHalf}
            />
          </section>
          <section aria-labelledby="account-limits-heading" className="flex flex-col gap-4">
            <h2 id="account-limits-heading" className="text-lg font-semibold">
              Usage limits
            </h2>
            <UsageLimitsSection selectedEnvironmentIds={null} now={limitsNow} />
          </section>
        </WorkspacePageContainer>
      </div>
    </>
  );
}

export const Route = createFileRoute("/account")({
  component: AccountRoute,
});
