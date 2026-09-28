import {
  AuthAccessWriteScope,
  type AgentWorkspacesSettings as AgentWorkspacesConfig,
  type EnvironmentId,
} from "@t3tools/contracts";
import { useAtomValue } from "@effect/atom-react";
import { BotIcon } from "lucide-react";
import { useState } from "react";

import { isElectron } from "../../env";
import { usePrimaryEnvironmentId } from "../../state/environments";
import { serverEnvironment } from "../../state/server";
import { useEnvironmentSessionState } from "../../state/session";
import { useAtomCommand } from "../../state/use-atom-command";
import { runtimeModeConfig, runtimeModeOptions } from "../chat/runtimeModeConfig";
import { Button } from "../ui/button";
import { Select, SelectItem, SelectPopup, SelectTrigger, SelectValue } from "../ui/select";
import { Textarea } from "../ui/textarea";
import { SettingsPageContainer, SettingsSection } from "./settingsLayout";
import { useSettingsScope } from "./SettingsScopeContext";
import { AGENT_WORKSPACE_OPTIONS, defaultAgentWorkspaceProfiles } from "../agentWorkspaces";

function AgentWorkspacesEditor({ environmentId }: { environmentId: EnvironmentId }) {
  const settings = useAtomValue(serverEnvironment.settingsValueAtom(environmentId));
  const session = useEnvironmentSessionState(environmentId);
  const primaryEnvironmentId = usePrimaryEnvironmentId();
  const canEdit =
    (isElectron && environmentId === primaryEnvironmentId) ||
    (session.data?.authenticated === true &&
      (session.data.scopes?.includes(AuthAccessWriteScope) ?? false));
  const saved = settings?.agentWorkspaces;
  const [draft, setDraft] = useState<AgentWorkspacesConfig | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const updateSettings = useAtomCommand(
    serverEnvironment.updateSettings,
    "agent workspaces update",
  );

  if (!saved) {
    return (
      <SettingsPageContainer width="wide" className="gap-6">
        <p className="px-3 text-sm text-muted-foreground sm:px-4">
          Connect to this server to configure agent workspaces.
        </p>
      </SettingsPageContainer>
    );
  }

  const profiles = draft ?? defaultAgentWorkspaceProfiles(saved);
  const changeProfile = (
    kind: keyof AgentWorkspacesConfig,
    patch: Partial<AgentWorkspacesConfig[keyof AgentWorkspacesConfig]>,
  ) => {
    setDraft((current) => {
      const next = current ?? defaultAgentWorkspaceProfiles(saved);
      return { ...next, [kind]: { ...next[kind], ...patch } };
    });
  };
  const save = async () => {
    if (!canEdit) return;
    setSaving(true);
    setError("");
    const next = defaultAgentWorkspaceProfiles(profiles);
    try {
      const result = await updateSettings({
        environmentId,
        input: { patch: { agentWorkspaces: next } },
      });
      if (result._tag === "Failure") throw new Error("Could not save agent workspaces.");
      setDraft(null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Could not save agent workspaces.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <SettingsPageContainer width="wide" className="gap-6">
      <div id="agent-workspaces" className="px-3 sm:px-4">
        <h1 className="text-lg font-medium">Agent workspaces</h1>
        <p className="mt-1 text-sm text-muted-foreground">
          Set access and extra instructions for each project workflow. File conversations stay in
          their open viewer.
        </p>
        {!canEdit ? (
          <p className="mt-2 text-sm text-muted-foreground">
            An admin session is required to change these settings.
          </p>
        ) : null}
      </div>

      {AGENT_WORKSPACE_OPTIONS.map(({ kind, name, description }) => {
        const profile = profiles[kind];
        return (
          <SettingsSection
            key={kind}
            title={name}
            icon={<BotIcon className="size-3.5 text-sidebar-muted-foreground/60" />}
          >
            <div className="space-y-4 px-3 py-3 sm:px-4">
              <p className="text-sm text-muted-foreground">{description}</p>
              {kind === "application" ? (
                <div className="max-w-sm space-y-1.5 text-sm">
                  <span>Tool access</span>
                  <p className="text-xs text-muted-foreground">
                    Supervised. Commands and file changes require approval, and the agent is told to
                    keep this workflow read-only.
                  </p>
                </div>
              ) : (
                <label className="block max-w-sm space-y-1.5 text-sm">
                  <span>Tool access</span>
                  <Select
                    value={profile.runtimeMode}
                    onValueChange={(value) => {
                      if (
                        value === "approval-required" ||
                        value === "auto-accept-edits" ||
                        value === "auto" ||
                        value === "full-access"
                      ) {
                        changeProfile(kind, { runtimeMode: value });
                      }
                    }}
                    disabled={!canEdit}
                  >
                    <SelectTrigger>
                      <SelectValue />
                    </SelectTrigger>
                    <SelectPopup>
                      {runtimeModeOptions.map((mode) => (
                        <SelectItem key={mode} value={mode}>
                          {runtimeModeConfig[mode].label}
                        </SelectItem>
                      ))}
                    </SelectPopup>
                  </Select>
                  <span className="block text-xs text-muted-foreground">
                    {runtimeModeConfig[profile.runtimeMode].description}
                  </span>
                </label>
              )}
              <label className="block space-y-1.5 text-sm">
                <span>Extra instructions</span>
                <Textarea
                  value={profile.instructions}
                  maxLength={4_000}
                  disabled={!canEdit}
                  placeholder="Add instructions for this workflow"
                  onChange={(event) => changeProfile(kind, { instructions: event.target.value })}
                />
              </label>
            </div>
          </SettingsSection>
        );
      })}

      {error ? (
        <p role="alert" className="px-3 text-sm text-destructive sm:px-4">
          {error}
        </p>
      ) : null}
      <div className="flex justify-end gap-2 px-3 sm:px-4">
        {draft ? (
          <Button variant="ghost" disabled={saving} onClick={() => setDraft(null)}>
            Discard
          </Button>
        ) : null}
        <Button disabled={!draft || !canEdit || saving} onClick={() => void save()}>
          {saving ? "Saving…" : "Save agent workspaces"}
        </Button>
      </div>
    </SettingsPageContainer>
  );
}

export function AgentWorkspacesSettingsPanel() {
  const { scope } = useSettingsScope();
  if (scope.kind !== "environment") {
    return (
      <SettingsPageContainer width="wide" className="gap-6">
        <div className="px-3 sm:px-4">
          <h1 className="text-lg font-medium">Agent workspaces</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Choose one server to configure its agent workspaces.
          </p>
        </div>
      </SettingsPageContainer>
    );
  }
  return <AgentWorkspacesEditor key={scope.environmentId} environmentId={scope.environmentId} />;
}
