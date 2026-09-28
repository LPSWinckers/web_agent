import { useEffect, useState } from "react";
import { useAtomValue } from "@effect/atom-react";
import {
  AuthAccessWriteScope,
  DEFAULT_IMAGE_BANK_SETTINGS,
  type EnvironmentId,
  type ImageBankColumn,
  type ImageBankDeviceCode,
  type ImageBankSettings as ImageBankSettingsValue,
} from "@t3tools/contracts";
import { ExternalLinkIcon } from "lucide-react";
import { isElectron } from "../../env";
import { usePrimaryEnvironmentId } from "../../state/environments";
import { useEnvironmentSessionState } from "../../state/session";
import { serverEnvironment } from "../../state/server";
import { useAtomCommand } from "../../state/use-atom-command";
import { squashAtomCommandFailure } from "@t3tools/client-runtime/state/runtime";
import { Button } from "../ui/button";
import { Input } from "../ui/input";

const MAPPING_FIELDS = [
  ["Themes", "themes"],
  ["Image type", "imageType"],
  ["Workfields", "workfields"],
  ["Keywords", "keywords"],
  ["Usage restrictions", "restrictions"],
] as const;

function messageOf(cause: unknown): string {
  return cause instanceof Error ? cause.message : "The image bank request failed.";
}

export function ImageBankSettings({ environmentId }: { environmentId: EnvironmentId }) {
  const settings = useAtomValue(serverEnvironment.settingsValueAtom(environmentId));
  const session = useEnvironmentSessionState(environmentId);
  const primaryEnvironmentId = usePrimaryEnvironmentId();
  const canEdit =
    (isElectron && environmentId === primaryEnvironmentId) ||
    (session.data?.authenticated === true &&
      (session.data.scopes?.includes(AuthAccessWriteScope) ?? false));
  const saved = settings?.imageBank ?? DEFAULT_IMAGE_BANK_SETTINGS;
  const [draft, setDraft] = useState<ImageBankSettingsValue>(saved);
  const [columns, setColumns] = useState<ReadonlyArray<ImageBankColumn>>([]);
  const [deviceCode, setDeviceCode] = useState<ImageBankDeviceCode | null>(null);
  const [status, setStatus] = useState<{ connected: boolean; signedInAs: string | null } | null>(
    null,
  );
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const updateSettings = useAtomCommand(
    serverEnvironment.updateSettings,
    "image bank settings save",
  );
  const getStatus = useAtomCommand(serverEnvironment.imageBankGetStatus, "image bank status");
  const startLogin = useAtomCommand(serverEnvironment.imageBankStartLogin, "image bank sign in");
  const pollLogin = useAtomCommand(serverEnvironment.imageBankPollLogin, "image bank sign in poll");
  const testConnection = useAtomCommand(
    serverEnvironment.imageBankTestConnection,
    "image bank test",
  );
  const disconnect = useAtomCommand(serverEnvironment.imageBankDisconnect, "image bank disconnect");

  useEffect(() => setDraft(saved), [saved]);
  useEffect(() => {
    let active = true;
    void getStatus({ environmentId, input: {} }).then((result) => {
      if (active && result._tag === "Success") setStatus(result.value);
    });
    return () => {
      active = false;
    };
  }, [environmentId, getStatus]);
  useEffect(() => {
    if (!deviceCode) return;
    let active = true;
    let polling = false;
    const timer = window.setInterval(
      () => {
        if (polling || !active) return;
        if (Date.now() >= deviceCode.expiresAt) {
          setDeviceCode(null);
          setError("The Microsoft sign-in code expired. Start sign-in again.");
          return;
        }
        polling = true;
        void pollLogin({ environmentId, input: {} })
          .then((result) => {
            if (!active) return;
            if (result._tag === "Failure") {
              setError(messageOf(squashAtomCommandFailure(result)));
              setDeviceCode(null);
            } else if (result.value.status === "connected") {
              setStatus({ connected: true, signedInAs: result.value.signedInAs });
              setDeviceCode(null);
              setMessage(`Connected as ${result.value.signedInAs}.`);
              setError("");
            } else if (result.value.status === "expired") {
              setDeviceCode(null);
              setError("The Microsoft sign-in code expired. Start sign-in again.");
            } else if (
              result.value.status === "pending" &&
              result.value.intervalSeconds !== undefined
            ) {
              const intervalSeconds = result.value.intervalSeconds;
              setDeviceCode((current) => (current ? { ...current, intervalSeconds } : null));
            }
          })
          .finally(() => {
            polling = false;
          });
      },
      Math.max(3, deviceCode.intervalSeconds) * 1000,
    );
    return () => {
      active = false;
      window.clearInterval(timer);
    };
  }, [deviceCode, environmentId, pollLogin]);

  const change = (patch: Partial<ImageBankSettingsValue>) =>
    setDraft((current) => ({ ...current, ...patch }));
  const save = async () => {
    if (!canEdit) return;
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const normalized = { ...draft, tenantId: draft.tenantId.trim() || "organizations" };
      const result = await updateSettings({
        environmentId,
        input: { patch: { imageBank: normalized } },
      });
      if (result._tag === "Failure") throw squashAtomCommandFailure(result);
      setDraft(normalized);
      const statusResult = await getStatus({ environmentId, input: {} });
      if (statusResult._tag === "Success") setStatus(statusResult.value);
      setMessage("Image-bank settings saved.");
    } catch (cause) {
      setError(messageOf(cause));
    } finally {
      setBusy(false);
    }
  };
  const beginLogin = async () => {
    setBusy(true);
    setError("");
    setMessage("");
    setDeviceCode(null);
    try {
      const normalized = { ...draft, tenantId: draft.tenantId.trim() || "organizations" };
      const saveResult = await updateSettings({
        environmentId,
        input: { patch: { imageBank: normalized } },
      });
      if (saveResult._tag === "Failure") throw squashAtomCommandFailure(saveResult);
      setDraft(normalized);
      const statusResult = await getStatus({ environmentId, input: {} });
      if (statusResult._tag === "Success") setStatus(statusResult.value);
      const result = await startLogin({ environmentId, input: {} });
      if (result._tag === "Failure") throw squashAtomCommandFailure(result);
      setDeviceCode(result.value);
    } catch (cause) {
      setError(messageOf(cause));
    } finally {
      setBusy(false);
    }
  };
  const test = async () => {
    setBusy(true);
    setError("");
    setMessage("");
    try {
      const result = await testConnection({ environmentId, input: {} });
      if (result._tag === "Failure") throw squashAtomCommandFailure(result);
      setColumns(result.value.columns);
      setMessage(`Connected to ${result.value.libraryName}. Map its metadata columns, then save.`);
    } catch (cause) {
      setError(messageOf(cause));
    } finally {
      setBusy(false);
    }
  };
  const signOut = async () => {
    setBusy(true);
    setError("");
    try {
      const result = await disconnect({ environmentId, input: {} });
      if (result._tag === "Failure") throw squashAtomCommandFailure(result);
      setStatus(result.value);
      setMessage("Microsoft sign-in removed from this T3 environment.");
    } catch (cause) {
      setError(messageOf(cause));
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className="space-y-4 px-3 sm:px-4" aria-labelledby="image-bank-title">
      <div>
        <h2 id="image-bank-title" className="text-sm font-medium">
          SharePoint image bank
        </h2>
        <p className="mt-1 max-w-3xl text-xs text-muted-foreground">
          Sign in with your Microsoft account. The agent can then search this library and import
          selected images into the current project. Access follows your SharePoint permissions.
        </p>
      </div>
      <div className="grid gap-3 sm:grid-cols-2">
        <label className="space-y-1.5 text-sm">
          <span>Microsoft tenant ID or domain</span>
          <Input
            value={draft.tenantId}
            maxLength={512}
            disabled={!canEdit || busy}
            onChange={(event) => change({ tenantId: event.target.value })}
            placeholder="organizations"
          />
        </label>
        <label className="space-y-1.5 text-sm">
          <span>Public-client application ID</span>
          <Input
            value={draft.clientId}
            maxLength={512}
            disabled={!canEdit || busy}
            onChange={(event) => change({ clientId: event.target.value })}
            placeholder="Microsoft Entra application (client) ID"
          />
        </label>
      </div>
      <label className="block space-y-1.5 text-sm">
        <span>SharePoint image-library link</span>
        <Input
          value={draft.shareUrl}
          maxLength={2048}
          disabled={!canEdit || busy}
          onChange={(event) => change({ shareUrl: event.target.value })}
          placeholder="https://...sharepoint.com/..."
        />
      </label>
      <p className="max-w-3xl text-xs text-muted-foreground">
        Use SharePoint’s Copy link for the library folder. The signed-in Microsoft account must
        already have access to it.
      </p>
      <p className="max-w-3xl text-xs text-muted-foreground">
        The Entra app must allow device-code sign-in and delegated Files.ReadWrite. The prototype
        only reads from SharePoint; imported copies go into the project workspace. No client secret
        or app-only access is used. If your tenant blocks user consent, its policy can still require
        an administrator to approve the delegated app.
      </p>
      {!canEdit ? (
        <p className="text-xs text-muted-foreground">
          An admin session for this T3 environment is required to change these settings.
        </p>
      ) : null}
      <div className="flex flex-wrap items-center gap-2">
        <Button size="sm" variant="outline" disabled={!canEdit || busy} onClick={() => void save()}>
          {busy ? "Saving…" : "Save settings"}
        </Button>
        <Button
          size="sm"
          disabled={!canEdit || busy || !draft.clientId.trim()}
          onClick={() => void beginLogin()}
        >
          {status?.connected ? "Sign in again" : "Sign in with Microsoft"}
        </Button>
        <Button
          size="sm"
          variant="outline"
          disabled={!canEdit || busy || !status?.connected}
          onClick={() => void test()}
        >
          Test library and load columns
        </Button>
        {status?.connected ? (
          <Button
            size="sm"
            variant="ghost"
            disabled={!canEdit || busy}
            onClick={() => void signOut()}
          >
            Disconnect
          </Button>
        ) : null}
        {status?.connected ? (
          <span className="text-xs text-muted-foreground">
            Connected as {status.signedInAs ?? "Microsoft account"}
          </span>
        ) : null}
      </div>
      {deviceCode ? (
        <div className="max-w-xl rounded-md border p-3 text-sm">
          <p>Open Microsoft sign-in and enter this code:</p>
          <a
            className="mt-2 inline-flex items-center gap-1 text-primary underline"
            href={deviceCode.verificationUri}
            target="_blank"
            rel="noreferrer"
          >
            {deviceCode.verificationUri}
            <ExternalLinkIcon className="size-3" />
          </a>
          <p className="mt-2 select-all font-mono text-lg font-semibold">{deviceCode.userCode}</p>
          <p className="mt-1 text-xs text-muted-foreground">
            This page checks sign-in every {deviceCode.intervalSeconds} seconds. The code expires in
            about {Math.max(1, Math.ceil((deviceCode.expiresAt - Date.now()) / 60_000))} minutes.
          </p>
        </div>
      ) : null}
      {columns.length > 0 ? (
        <>
          <div className="grid gap-3 sm:grid-cols-2">
            {MAPPING_FIELDS.map(([label, key]) => (
              <label className="space-y-1.5 text-sm" key={key}>
                <span>{label} column</span>
                <select
                  className="w-full rounded-md border bg-background p-2 text-sm"
                  value={draft.fieldMapping[key]}
                  disabled={!canEdit || busy}
                  onChange={(event) =>
                    change({ fieldMapping: { ...draft.fieldMapping, [key]: event.target.value } })
                  }
                >
                  <option value="">Not mapped</option>
                  {columns.map((column) => (
                    <option value={column.name} key={column.name}>
                      {column.displayName}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
          <p className="max-w-3xl text-xs text-muted-foreground">
            If restriction tags are stored with keywords, map the same column to both. Tags such as
            “NIET VOOR ADVERTING” are treated as restrictions; ordinary keywords are not.
          </p>
        </>
      ) : null}
      {error ? (
        <p role="alert" className="text-sm text-destructive">
          {error}
        </p>
      ) : null}
      {message ? (
        <p role="status" className="text-sm text-muted-foreground">
          {message}
        </p>
      ) : null}
    </section>
  );
}
