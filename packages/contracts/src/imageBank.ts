import * as Schema from "effect/Schema";

import { NonNegativeInt, TrimmedNonEmptyString, TrimmedString } from "./baseSchemas.ts";

const IMAGE_BANK_TEXT_LIMIT = 512;

export const ImageBankFieldMapping = Schema.Struct({
  themes: TrimmedString.check(Schema.isMaxLength(100)),
  imageType: TrimmedString.check(Schema.isMaxLength(100)),
  workfields: TrimmedString.check(Schema.isMaxLength(100)),
  keywords: TrimmedString.check(Schema.isMaxLength(100)),
  restrictions: TrimmedString.check(Schema.isMaxLength(100)),
});
export type ImageBankFieldMapping = typeof ImageBankFieldMapping.Type;

export const EMPTY_IMAGE_BANK_FIELD_MAPPING: ImageBankFieldMapping = {
  themes: "",
  imageType: "",
  workfields: "",
  keywords: "",
  restrictions: "",
};

/** Tenant, public-client app and shared image-folder details. Tokens live in the secret store. */
export const ImageBankSettings = Schema.Struct({
  tenantId: TrimmedString.check(Schema.isMaxLength(IMAGE_BANK_TEXT_LIMIT)),
  clientId: TrimmedString.check(Schema.isMaxLength(IMAGE_BANK_TEXT_LIMIT)),
  shareUrl: TrimmedString.check(Schema.isMaxLength(2048)),
  fieldMapping: ImageBankFieldMapping,
});
export type ImageBankSettings = typeof ImageBankSettings.Type;

export const DEFAULT_IMAGE_BANK_SETTINGS: ImageBankSettings = {
  tenantId: "organizations",
  clientId: "",
  shareUrl: "",
  fieldMapping: EMPTY_IMAGE_BANK_FIELD_MAPPING,
};

export const ImageBankConnectionStatus = Schema.Struct({
  configured: Schema.Boolean,
  connected: Schema.Boolean,
  signedInAs: Schema.NullOr(Schema.String),
});
export type ImageBankConnectionStatus = typeof ImageBankConnectionStatus.Type;

export const ImageBankDeviceCode = Schema.Struct({
  verificationUri: Schema.String,
  userCode: Schema.String,
  message: Schema.String,
  expiresAt: NonNegativeInt,
  intervalSeconds: NonNegativeInt,
});
export type ImageBankDeviceCode = typeof ImageBankDeviceCode.Type;

export const ImageBankLoginPollResult = Schema.Union([
  Schema.Struct({
    status: Schema.Literal("pending"),
    intervalSeconds: Schema.optional(NonNegativeInt),
  }),
  Schema.Struct({ status: Schema.Literal("connected"), signedInAs: Schema.String }),
  Schema.Struct({ status: Schema.Literal("expired") }),
]);
export type ImageBankLoginPollResult = typeof ImageBankLoginPollResult.Type;

export const ImageBankColumn = Schema.Struct({
  name: TrimmedNonEmptyString,
  displayName: TrimmedNonEmptyString,
});
export type ImageBankColumn = typeof ImageBankColumn.Type;

export const ImageBankConnectionTestResult = Schema.Struct({
  libraryName: Schema.String,
  columns: Schema.Array(ImageBankColumn),
});
export type ImageBankConnectionTestResult = typeof ImageBankConnectionTestResult.Type;

export class ImageBankError extends Schema.TaggedError<ImageBankError>()("ImageBankError", {
  operation: Schema.Literals(["configure", "sign-in", "test", "search", "import", "disconnect"]),
  detail: Schema.String,
}) {
  override get message(): string {
    return this.detail;
  }
}

export const ImageBankSearchInput = Schema.Struct({
  query: TrimmedNonEmptyString.check(Schema.isMaxLength(500)),
  themes: Schema.optional(TrimmedString.check(Schema.isMaxLength(200))),
  workfields: Schema.optional(TrimmedString.check(Schema.isMaxLength(200))),
  imageType: Schema.optional(TrimmedString.check(Schema.isMaxLength(100))),
  limit: Schema.optional(Schema.Int.check(Schema.isBetween({ minimum: 1, maximum: 12 }))),
});
export type ImageBankSearchInput = typeof ImageBankSearchInput.Type;

export const ImageBankCandidate = Schema.Struct({
  assetId: TrimmedNonEmptyString,
  name: TrimmedNonEmptyString,
  fileName: TrimmedNonEmptyString,
  webUrl: Schema.String,
  mimeType: Schema.String,
  sizeBytes: NonNegativeInt,
  themes: Schema.Array(Schema.String),
  imageType: Schema.Array(Schema.String),
  workfields: Schema.Array(Schema.String),
  keywords: Schema.Array(Schema.String),
  restrictions: Schema.Array(Schema.String),
  restricted: Schema.Boolean,
});
export type ImageBankCandidate = typeof ImageBankCandidate.Type;

export const ImageBankSearchResult = Schema.Struct({
  candidates: Schema.Array(ImageBankCandidate),
});
export type ImageBankSearchResult = typeof ImageBankSearchResult.Type;

export const ImageBankImportInput = Schema.Struct({
  assetId: TrimmedNonEmptyString.check(Schema.isMaxLength(512)),
  relativePath: Schema.optional(TrimmedNonEmptyString.check(Schema.isMaxLength(1024))),
  allowRestricted: Schema.optional(Schema.Boolean),
});
export type ImageBankImportInput = typeof ImageBankImportInput.Type;

export const ImageBankImportedAsset = Schema.Struct({
  relativePath: TrimmedNonEmptyString,
  fileName: TrimmedNonEmptyString,
  mimeType: TrimmedNonEmptyString,
  sourceItemId: TrimmedNonEmptyString,
  sourceUrl: Schema.String,
  themes: Schema.Array(Schema.String),
  imageType: Schema.Array(Schema.String),
  workfields: Schema.Array(Schema.String),
  keywords: Schema.Array(Schema.String),
  restrictions: Schema.Array(Schema.String),
});
export type ImageBankImportedAsset = typeof ImageBankImportedAsset.Type;
