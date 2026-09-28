import {
  ImageBankError,
  ImageBankSearchInput,
  ImageBankSearchResult,
  ImageBankImportInput,
  ImageBankImportedAsset,
  McpCapabilityUnavailableError,
} from "@t3tools/contracts";
import * as Schema from "effect/Schema";
import * as Tool from "effect/unstable/ai/Tool";
import * as Toolkit from "effect/unstable/ai/Toolkit";
import * as ImageBankService from "../../../imageBank/ImageBankService.ts";
import * as McpInvocationContext from "../../McpInvocationContext.ts";
import * as ProjectionSnapshotQuery from "../../../orchestration/Services/ProjectionSnapshotQuery.ts";
import * as WorkspaceFileSystem from "../../../workspace/WorkspaceFileSystem.ts";

const dependencies = [
  ImageBankService.ImageBankService,
  McpInvocationContext.McpInvocationContext,
  ProjectionSnapshotQuery.ProjectionSnapshotQuery,
  WorkspaceFileSystem.WorkspaceFileSystem,
];

const ImageBankToolError = Schema.Union([ImageBankError, McpCapabilityUnavailableError]);

const ImageBankSearchTool = Tool.make("image_bank_search", {
  description:
    "Search the connected SharePoint image bank by filename, themes, image type, workfields, or keywords. Results include usage restrictions. Exclude restricted results by default, especially items tagged NIET VOOR ADVERTING. Search results never include image binaries.",
  parameters: ImageBankSearchInput,
  success: ImageBankSearchResult,
  failure: ImageBankToolError,
  dependencies,
})
  .annotate(Tool.Title, "Search SharePoint image bank")
  .annotate(Tool.Readonly, true)
  .annotate(Tool.Destructive, false)
  .annotate(Tool.Idempotent, true)
  .annotate(Tool.OpenWorld, true);

const ImageBankImportTool = Tool.make("image_bank_import", {
  description:
    "Import one selected SharePoint image into the current project workspace. Search first and choose the closest fit from its metadata. Images with any usage restriction are blocked unless the user explicitly confirms permission with allowRestricted=true. Returns the project-relative path and source metadata for a presentation image reference.",
  parameters: ImageBankImportInput,
  success: ImageBankImportedAsset,
  failure: ImageBankToolError,
  dependencies,
})
  .annotate(Tool.Title, "Import SharePoint image")
  .annotate(Tool.Readonly, false)
  .annotate(Tool.Destructive, false)
  .annotate(Tool.Idempotent, true)
  .annotate(Tool.OpenWorld, true);

export const ImageBankToolkit = Toolkit.make(ImageBankSearchTool, ImageBankImportTool);
