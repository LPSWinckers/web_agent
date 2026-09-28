import { useState } from "react";
import {
  AuthAccessWriteScope,
  type CompanyLibrary,
  type CompanyLibraryEntry,
  type EnvironmentId,
} from "@t3tools/contracts";
import { useAtomValue } from "@effect/atom-react";
import { PlusIcon, Trash2Icon } from "lucide-react";

import { isElectron } from "../../env";
import { randomUUID } from "../../lib/utils";
import { usePrimaryEnvironmentId } from "../../state/environments";
import { useEnvironmentSessionState } from "../../state/session";
import { serverEnvironment } from "../../state/server";
import { useAtomCommand } from "../../state/use-atom-command";
import { Button } from "../ui/button";
import { Input } from "../ui/input";
import { Textarea } from "../ui/textarea";
import { SettingsPageContainer } from "./settingsLayout";
import { useSettingsScope } from "./SettingsScopeContext";
import { ConsultancyStandardGallery } from "./ConsultancyStandardGallery";
import { ImageBankSettings } from "./ImageBankSettings";

type EntryGroup = "skills";

const GROUPS: ReadonlyArray<{
  key: EntryGroup;
  title: string;
  description: string;
  singular: string;
}> = [
  {
    key: "skills",
    title: "Skills",
    description: "Reusable procedures agents can follow when the request names the skill.",
    singular: "skill",
  },
];

function validLibrary(library: CompanyLibrary) {
  const word = library.wordStandard;
  return (
    library.companyName.length <= 100 &&
    library.guidance.length <= 3000 &&
    word.name.trim().length > 0 &&
    word.name.length <= 100 &&
    word.fontFamily.trim().length > 0 &&
    word.fontFamily.length <= 100 &&
    [word.bodyColor, word.headingColor, word.accentColor].every((color) =>
      /^[0-9A-Fa-f]{6}$/.test(color),
    ) &&
    word.sections.filter((section) => section.trim()).length >= 1 &&
    word.sections.filter((section) => section.trim()).length <= 12 &&
    word.sections.every((section) => section.length <= 100) &&
    GROUPS.every(
      ({ key }) =>
        library[key].length <= 20 &&
        library[key].every(
          (entry) =>
            entry.name.trim().length > 0 &&
            entry.name.length <= 100 &&
            entry.instructions.trim().length > 0 &&
            entry.instructions.length <= 2000,
        ),
    )
  );
}

function CompanySettingsEditor({ environmentId }: { environmentId: EnvironmentId }) {
  const settings = useAtomValue(serverEnvironment.settingsValueAtom(environmentId));
  const session = useEnvironmentSessionState(environmentId);
  const primaryEnvironmentId = usePrimaryEnvironmentId();
  const canEdit =
    (isElectron && environmentId === primaryEnvironmentId) ||
    (session.data?.authenticated === true &&
      (session.data.scopes?.includes(AuthAccessWriteScope) ?? false));
  const saved = settings?.companyLibrary;
  const [editedDraft, setEditedDraft] = useState<CompanyLibrary | null>(null);
  const draft = editedDraft ?? saved;
  const editing = editedDraft !== null;
  const [saving, setSaving] = useState(false);
  const updateSettings = useAtomCommand(
    serverEnvironment.updateSettings,
    "company settings update",
  );

  if (!saved || !draft) {
    return (
      <SettingsPageContainer width="wide" className="gap-6">
        <ConsultancyStandardGallery />
        <p className="px-3 text-sm text-muted-foreground sm:px-4">
          Connect to this server to view its company library.
        </p>
      </SettingsPageContainer>
    );
  }

  const change = (patch: Partial<CompanyLibrary>) => {
    setEditedDraft({ ...draft, ...patch });
  };
  const changeEntry = (key: EntryGroup, id: string, patch: Partial<CompanyLibraryEntry>) => {
    change({
      [key]: draft[key].map((entry) => (entry.id === id ? { ...entry, ...patch } : entry)),
    });
  };
  const addEntry = (key: EntryGroup) => {
    if (draft[key].length >= 20) return;
    change({ [key]: [...draft[key], { id: randomUUID(), name: "", instructions: "" }] });
  };
  const save = async () => {
    if (!canEdit || !validLibrary(draft)) return;
    setSaving(true);
    try {
      const normalized: CompanyLibrary = {
        companyName: draft.companyName.trim(),
        guidance: draft.guidance.trim(),
        chartTemplates: draft.chartTemplates,
        skills: draft.skills.map((entry) => ({
          ...entry,
          name: entry.name.trim(),
          instructions: entry.instructions.trim(),
        })),
        powerpointStandards: draft.powerpointStandards,
        wordStandard: {
          ...draft.wordStandard,
          name: draft.wordStandard.name.trim(),
          fontFamily: draft.wordStandard.fontFamily.trim(),
          sections: draft.wordStandard.sections.map((section) => section.trim()).filter(Boolean),
        },
      };
      const result = await updateSettings({
        environmentId,
        input: { patch: { companyLibrary: normalized } },
      });
      if (result._tag !== "Failure") setEditedDraft(null);
    } finally {
      setSaving(false);
    }
  };

  return (
    <SettingsPageContainer width="wide" className="gap-6">
      <div className="px-3 sm:px-4">
        <h1 className="text-lg font-medium">Company library</h1>
        <p className="mt-1 text-sm text-muted-foreground">
          Word document settings, company guidance, and skills belong to this server. The
          presentation and chart standard is shared by everyone.
        </p>
        {!canEdit ? (
          <p className="mt-2 text-sm text-muted-foreground">
            An admin session is required to edit the Word standard, company guidance, and skills.
          </p>
        ) : null}
      </div>

      <ImageBankSettings environmentId={environmentId} />

      <ConsultancyStandardGallery />

      <section className="space-y-4 px-3 sm:px-4" aria-labelledby="word-standard-title">
        <div>
          <h2 id="word-standard-title" className="text-sm font-medium">
            Word document standard
          </h2>
          <p className="mt-1 text-xs text-muted-foreground">
            New documents use this theme and section order. Apply it to an open document from the
            editor.
          </p>
        </div>
        <div className="grid gap-3 sm:grid-cols-2">
          <label className="space-y-1.5 text-sm">
            <span>Standard name</span>
            <Input
              value={draft.wordStandard.name}
              maxLength={100}
              disabled={!canEdit}
              onChange={(event) =>
                change({ wordStandard: { ...draft.wordStandard, name: event.target.value } })
              }
            />
          </label>
          <label className="space-y-1.5 text-sm">
            <span>Font family</span>
            <Input
              value={draft.wordStandard.fontFamily}
              maxLength={100}
              disabled={!canEdit}
              onChange={(event) =>
                change({ wordStandard: { ...draft.wordStandard, fontFamily: event.target.value } })
              }
            />
          </label>
          {(
            [
              ["Body color", "bodyColor"],
              ["Heading color", "headingColor"],
              ["Accent color", "accentColor"],
            ] as const
          ).map(([label, key]) => (
            <label key={key} className="space-y-1.5 text-sm">
              <span>{label} (hex)</span>
              <div className="flex items-center gap-2">
                <span
                  className="size-7 shrink-0 rounded border"
                  style={{
                    backgroundColor: /^[0-9A-Fa-f]{6}$/.test(draft.wordStandard[key])
                      ? `#${draft.wordStandard[key]}`
                      : undefined,
                  }}
                />
                <Input
                  aria-label={label}
                  value={draft.wordStandard[key]}
                  maxLength={6}
                  disabled={!canEdit}
                  onChange={(event) =>
                    change({ wordStandard: { ...draft.wordStandard, [key]: event.target.value } })
                  }
                />
              </div>
            </label>
          ))}
        </div>
        <label className="block space-y-1.5 text-sm">
          <span>Section headings, one per line</span>
          <Textarea
            value={draft.wordStandard.sections.join("\n")}
            disabled={!canEdit}
            onChange={(event) =>
              change({
                wordStandard: { ...draft.wordStandard, sections: event.target.value.split("\n") },
              })
            }
          />
        </label>
        <div
          className="max-w-lg border border-border bg-white px-7 py-8 shadow-sm"
          style={{
            fontFamily: draft.wordStandard.fontFamily || "Arial",
            color: `#${/^[0-9A-Fa-f]{6}$/.test(draft.wordStandard.bodyColor) ? draft.wordStandard.bodyColor : "263445"}`,
          }}
          aria-label="Word standard preview"
        >
          <h3
            className="text-2xl font-bold"
            style={{
              color: `#${/^[0-9A-Fa-f]{6}$/.test(draft.wordStandard.headingColor) ? draft.wordStandard.headingColor : "1F345E"}`,
            }}
          >
            Voorbeeldrapport
          </h3>
          <h4
            className="mt-7 border-b-2 pb-1 text-lg font-bold"
            style={{
              borderColor: `#${/^[0-9A-Fa-f]{6}$/.test(draft.wordStandard.accentColor) ? draft.wordStandard.accentColor : "0075AB"}`,
            }}
          >
            {draft.wordStandard.sections.find((section) => section.trim()) ?? "Sectie"}
          </h4>
          <p className="mt-3 text-sm leading-relaxed">
            De opmaak en volgorde gelden voor nieuwe Word-documenten.
          </p>
        </div>
      </section>

      <section className="space-y-3 px-3 sm:px-4" aria-labelledby="company-settings-title">
        <h2 id="company-settings-title" className="text-sm text-foreground/70">
          Company settings
        </h2>
        <label className="block space-y-1.5 text-sm">
          <span>Company name</span>
          <Input
            value={draft.companyName}
            maxLength={100}
            disabled={!canEdit}
            onChange={(event) => change({ companyName: event.target.value })}
          />
        </label>
        <label className="block space-y-1.5 text-sm">
          <span>General agent guidance</span>
          <Textarea
            value={draft.guidance}
            maxLength={3000}
            disabled={!canEdit}
            onChange={(event) => change({ guidance: event.target.value })}
          />
        </label>
      </section>

      {GROUPS.map((group) => (
        <section
          key={group.key}
          className="space-y-3 px-3 sm:px-4"
          aria-labelledby={`${group.key}-title`}
        >
          <div className="flex items-start justify-between gap-3">
            <div>
              <h2 id={`${group.key}-title`} className="text-sm text-foreground/70">
                {group.title}
              </h2>
              <p className="mt-1 text-xs text-muted-foreground">{group.description}</p>
            </div>
            {canEdit ? (
              <Button
                size="xs"
                variant="outline"
                disabled={draft[group.key].length >= 20}
                onClick={() => addEntry(group.key)}
              >
                <PlusIcon className="size-3.5" /> Add
              </Button>
            ) : null}
          </div>
          {draft[group.key].length === 0 ? (
            <p className="rounded-xl border border-dashed border-border/70 px-4 py-5 text-sm text-muted-foreground">
              No {group.title.toLowerCase()} yet.
            </p>
          ) : (
            <div className="space-y-3">
              {draft[group.key].map((entry) => (
                <div
                  key={entry.id}
                  className="space-y-3 rounded-xl border border-border/60 bg-card/40 p-4"
                >
                  <div className="flex items-end gap-2">
                    <label className="min-w-0 flex-1 space-y-1.5 text-sm">
                      <span>Name</span>
                      <Input
                        value={entry.name}
                        maxLength={100}
                        disabled={!canEdit}
                        aria-label={`${group.singular} name`}
                        onChange={(event) =>
                          changeEntry(group.key, entry.id, { name: event.target.value })
                        }
                      />
                    </label>
                    {canEdit ? (
                      <Button
                        size="icon-xs"
                        variant="ghost-muted"
                        aria-label={`Remove ${entry.name || group.singular}`}
                        onClick={() =>
                          change({
                            [group.key]: draft[group.key].filter((item) => item.id !== entry.id),
                          })
                        }
                      >
                        <Trash2Icon className="size-3.5" />
                      </Button>
                    ) : null}
                  </div>
                  <label className="block space-y-1.5 text-sm">
                    <span>Instructions</span>
                    <Textarea
                      value={entry.instructions}
                      maxLength={2000}
                      disabled={!canEdit}
                      aria-label={`${group.singular} instructions`}
                      onChange={(event) =>
                        changeEntry(group.key, entry.id, { instructions: event.target.value })
                      }
                    />
                  </label>
                </div>
              ))}
            </div>
          )}
        </section>
      ))}

      {canEdit ? (
        <div className="flex justify-end gap-2 px-3 pb-8 sm:px-4">
          <Button
            size="sm"
            variant="outline"
            disabled={!editing || saving}
            onClick={() => {
              setEditedDraft(null);
            }}
          >
            Discard
          </Button>
          <Button
            size="sm"
            disabled={!editing || !validLibrary(draft) || saving}
            onClick={() => void save()}
          >
            {saving ? "Saving…" : "Save company library"}
          </Button>
        </div>
      ) : null}
    </SettingsPageContainer>
  );
}

export function CompanySettingsPanel() {
  const { scope } = useSettingsScope();
  if (scope.kind !== "environment") {
    return (
      <SettingsPageContainer width="wide" className="gap-6">
        <ConsultancyStandardGallery />
        <p className="px-3 text-sm text-muted-foreground sm:px-4">
          Choose one server to edit its company name, guidance, and skills.
        </p>
      </SettingsPageContainer>
    );
  }
  return <CompanySettingsEditor key={scope.environmentId} environmentId={scope.environmentId} />;
}
