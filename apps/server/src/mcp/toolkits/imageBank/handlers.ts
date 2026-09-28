import { ImageBankError } from "@t3tools/contracts";
import * as Effect from "effect/Effect";
import * as Layer from "effect/Layer";
import * as Option from "effect/Option";
import * as ImageBankService from "../../../imageBank/ImageBankService.ts";
import * as ProjectionSnapshotQuery from "../../../orchestration/Services/ProjectionSnapshotQuery.ts";
import * as WorkspaceFileSystem from "../../../workspace/WorkspaceFileSystem.ts";
import * as McpInvocationContext from "../../McpInvocationContext.ts";
import { ImageBankToolkit } from "./tools.ts";

const make = Effect.gen(function* () {
  const imageBank = yield* ImageBankService.ImageBankService;
  const snapshots = yield* ProjectionSnapshotQuery.ProjectionSnapshotQuery;
  const workspace = yield* WorkspaceFileSystem.WorkspaceFileSystem;

  return ImageBankToolkit.of({
    image_bank_search: (input) =>
      Effect.gen(function* () {
        yield* McpInvocationContext.requireMcpCapability("image-bank");
        return yield* imageBank.search(input);
      }),
    image_bank_import: (input) =>
      Effect.gen(function* () {
        const scope = yield* McpInvocationContext.requireMcpCapability("image-bank");
        const thread = yield* snapshots
          .getThreadShellById(scope.threadId)
          .pipe(
            Effect.mapError(
              () =>
                new ImageBankError({
                  operation: "import",
                  detail: "Could not find the current project workspace.",
                }),
            ),
          );
        if (Option.isNone(thread)) {
          return yield* new ImageBankError({
            operation: "import",
            detail: "Could not find the current project workspace.",
          });
        }
        const project = yield* snapshots
          .getProjectShellById(thread.value.projectId)
          .pipe(
            Effect.mapError(
              () =>
                new ImageBankError({
                  operation: "import",
                  detail: "Could not find the current project workspace.",
                }),
            ),
          );
        if (Option.isNone(project)) {
          return yield* new ImageBankError({
            operation: "import",
            detail: "Could not find the current project workspace.",
          });
        }
        const cwd = thread.value.worktreePath ?? project.value.workspaceRoot;
        const payload = yield* imageBank.import(input);
        yield* workspace
          .writeFile({
            cwd,
            relativePath: payload.asset.relativePath,
            contents: payload.contentsBase64,
            encoding: "base64",
          })
          .pipe(
            Effect.mapError(
              () =>
                new ImageBankError({
                  operation: "import",
                  detail: "Could not write the image into the project workspace.",
                }),
            ),
          );
        return payload.asset;
      }),
  });
});

export const ImageBankToolkitHandlersLive = ImageBankToolkit.toLayer(make).pipe(
  Layer.provide(ImageBankService.layer),
);
