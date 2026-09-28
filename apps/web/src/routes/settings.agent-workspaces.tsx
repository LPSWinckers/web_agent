import { createFileRoute } from "@tanstack/react-router";

import { AgentWorkspacesSettingsPanel } from "../components/settings/AgentWorkspacesSettings";

export const Route = createFileRoute("/settings/agent-workspaces")({
  component: AgentWorkspacesSettingsPanel,
});
