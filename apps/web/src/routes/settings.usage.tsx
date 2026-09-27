import { createFileRoute } from "@tanstack/react-router";

import { UsagePage } from "../components/usage/UsagePage";

function SettingsUsageRoute() {
  return <UsagePage settingsPage />;
}

export const Route = createFileRoute("/settings/usage")({
  component: SettingsUsageRoute,
});
