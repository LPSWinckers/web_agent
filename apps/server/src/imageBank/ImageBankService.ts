import * as Clock from "effect/Clock";
import * as Context from "effect/Context";
import * as Effect from "effect/Effect";
import * as Layer from "effect/Layer";
import * as Option from "effect/Option";
import * as Ref from "effect/Ref";
import * as Result from "effect/Result";
import * as Schema from "effect/Schema";
import * as HttpClient from "effect/unstable/http/HttpClient";
import * as HttpClientRequest from "effect/unstable/http/HttpClientRequest";
import {
  ImageBankError,
  type ImageBankCandidate,
  type ImageBankColumn,
  type ImageBankConnectionStatus,
  type ImageBankConnectionTestResult,
  type ImageBankDeviceCode,
  type ImageBankFieldMapping,
  type ImageBankImportedAsset,
  type ImageBankImportInput,
  type ImageBankLoginPollResult,
  type ImageBankSettings,
  type ImageBankSearchInput,
  type ImageBankSearchResult,
} from "@t3tools/contracts";
import * as ServerSecretStore from "../auth/ServerSecretStore.ts";
import * as ServerSettings from "../serverSettings.ts";

const TOKEN_SECRET = "sharepoint-image-bank-token";
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
// The Shares API requires this delegated scope to resolve a library sharing link.
// The prototype only sends read requests to Graph.
const TOKEN_SCOPES = "Files.ReadWrite offline_access User.Read";
const MAX_IMAGE_BYTES = 18 * 1024 * 1024;
const MAX_LIBRARY_ITEMS = 10_000;
const CACHE_TTL_MS = 5 * 60_000;

const TokenBundle = Schema.Struct({
  tenantId: Schema.String,
  clientId: Schema.String,
  refreshToken: Schema.String,
  signedInAs: Schema.String,
});
type TokenBundle = typeof TokenBundle.Type;
const TokenBundleJson = Schema.fromJsonString(TokenBundle);
const OAuthErrorResponse = Schema.Struct({
  error: Schema.String,
  error_description: Schema.optional(Schema.String),
});
const OAuthTokenResponse = Schema.Struct({
  access_token: Schema.String,
  refresh_token: Schema.optional(Schema.String),
  expires_in: Schema.optional(Schema.Number),
});
const DeviceCodeResponse = Schema.Struct({
  device_code: Schema.String,
  user_code: Schema.String,
  verification_uri: Schema.String,
  message: Schema.optional(Schema.String),
  expires_in: Schema.Number,
  interval: Schema.optional(Schema.Number),
});
const GraphErrorResponse = Schema.Struct({
  error: Schema.optional(Schema.Struct({ message: Schema.optional(Schema.String) })),
});
const GraphDriveItem = Schema.Struct({
  id: Schema.optional(Schema.String),
  name: Schema.optional(Schema.String),
  webUrl: Schema.optional(Schema.String),
  size: Schema.optional(Schema.Number),
  file: Schema.optional(Schema.Struct({ mimeType: Schema.optional(Schema.String) })),
  folder: Schema.optional(Schema.Unknown),
  parentReference: Schema.optional(Schema.Struct({ driveId: Schema.optional(Schema.String) })),
  listItem: Schema.optional(
    Schema.Struct({ fields: Schema.optional(Schema.Record(Schema.String, Schema.Unknown)) }),
  ),
});
const GraphResponse = Schema.Struct({
  value: Schema.optional(Schema.Array(GraphDriveItem)),
  "@odata.nextLink": Schema.optional(Schema.String),
});
const GraphProfile = Schema.Struct({
  displayName: Schema.optional(Schema.String),
  mail: Schema.optional(Schema.String),
  userPrincipalName: Schema.optional(Schema.String),
});

type GraphDriveItem = typeof GraphDriveItem.Type;
type GraphResponse = typeof GraphResponse.Type;

interface PendingDeviceLogin {
  readonly deviceCode: string;
  readonly tenantId: string;
  readonly clientId: string;
  readonly expiresAt: number;
  readonly intervalSeconds: number;
}

interface CachedItem {
  readonly candidate: ImageBankCandidate;
  readonly driveId: string;
  readonly itemId: string;
}

interface SearchCache {
  readonly expiresAt: number;
  readonly key: string;
  readonly items: ReadonlyArray<CachedItem>;
}

export interface ImageBankImportPayload {
  readonly asset: ImageBankImportedAsset;
  readonly contentsBase64: string;
}

const imageBankError = (operation: ImageBankError["operation"], detail: string) =>
  new ImageBankError({ operation, detail });

const describeError = (cause: unknown) =>
  cause instanceof Error ? cause.message : "The Microsoft request failed.";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringValue(value: unknown): string {
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return value.map(stringValue).filter(Boolean).join(", ");
  return "";
}

function valuesOf(value: unknown): string[] {
  return stringValue(value)
    .split(/[;,\n]/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function shareId(shareUrl: string): string {
  return `u!${Buffer.from(shareUrl, "utf8").toString("base64url")}`;
}

function displayFieldName(name: string): string {
  return name
    .replace(/_x([0-9a-f]{4})_/gi, (_match, code: string) =>
      String.fromCharCode(Number.parseInt(code, 16)),
    )
    .replace(/[_]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function safeFileName(fileName: string): string {
  const name = fileName
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-")
    .replace(/[. ]+$/, "")
    .slice(0, 160);
  return name || "sharepoint-image.jpg";
}

function mappedValues(fields: Record<string, unknown> | undefined, key: string): string[] {
  return key && fields ? valuesOf(fields[key]) : [];
}

function toCandidate(
  item: GraphDriveItem,
  mapping: ImageBankFieldMapping,
): ImageBankCandidate | null {
  const itemId = item.id;
  const fileName = item.name;
  const mimeType = item.file?.mimeType ?? "";
  if (!itemId || !fileName || !mimeType.startsWith("image/")) return null;
  if (
    !new Set(["image/jpeg", "image/png", "image/gif", "image/bmp", "image/tiff"]).has(
      mimeType.toLowerCase(),
    )
  ) {
    return null;
  }
  const fields = item.listItem?.fields;
  const keywords = mappedValues(fields, mapping.keywords);
  const themes = mappedValues(fields, mapping.themes);
  const imageType = mappedValues(fields, mapping.imageType);
  const workfields = mappedValues(fields, mapping.workfields);
  const classificationValues = [...themes, ...imageType, ...workfields, ...keywords];
  const mappedRestrictions = mappedValues(fields, mapping.restrictions);
  const restrictionsShareAClassificationColumn = Boolean(
    mapping.restrictions &&
    [mapping.themes, mapping.imageType, mapping.workfields, mapping.keywords].includes(
      mapping.restrictions,
    ),
  );
  const restrictions = [
    ...mappedRestrictions,
    ...classificationValues.filter(isRestrictionValue),
  ].filter(
    (value, index, values) =>
      (!restrictionsShareAClassificationColumn || isRestrictionValue(value)) &&
      values.findIndex((other) => normalize(other) === normalize(value)) === index,
  );
  return {
    assetId: itemId,
    name: fileName.replace(/\.[^.]+$/, ""),
    fileName,
    webUrl: item.webUrl ?? "",
    mimeType,
    sizeBytes: Math.max(0, Math.trunc(item.size ?? 0)),
    themes,
    imageType,
    workfields,
    keywords,
    restrictions,
    restricted: restrictions.length > 0,
  };
}

const normalize = (value: string) =>
  value
    .toLocaleLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

const isRestrictionValue = (value: string) =>
  /\b(niet\s+voor|not\s+for|do\s+not\s+use|restricted|verboden|alleen\s+intern|uitsluitend\s+intern)\b/.test(
    normalize(value),
  );

export class ImageBankService extends Context.Service<
  ImageBankService,
  {
    readonly getStatus: Effect.Effect<ImageBankConnectionStatus, ImageBankError>;
    readonly startLogin: Effect.Effect<ImageBankDeviceCode, ImageBankError>;
    readonly pollLogin: Effect.Effect<ImageBankLoginPollResult, ImageBankError>;
    readonly testConnection: Effect.Effect<ImageBankConnectionTestResult, ImageBankError>;
    readonly disconnect: Effect.Effect<ImageBankConnectionStatus, ImageBankError>;
    readonly search: (
      input: ImageBankSearchInput,
    ) => Effect.Effect<ImageBankSearchResult, ImageBankError>;
    readonly import: (
      input: ImageBankImportInput,
    ) => Effect.Effect<ImageBankImportPayload, ImageBankError>;
  }
>()("t3/imageBank/ImageBankService") {}

const make = Effect.gen(function* () {
  const settingsService = yield* ServerSettings.ServerSettingsService;
  const secretStore = yield* ServerSecretStore.ServerSecretStore;
  const httpClient = HttpClient.withScope(yield* HttpClient.HttpClient);
  const pendingLogin = yield* Ref.make<PendingDeviceLogin | null>(null);
  const accessToken = yield* Ref.make<{
    readonly token: string;
    readonly expiresAt: number;
  } | null>(null);
  const searchCache = yield* Ref.make<SearchCache | null>(null);

  const settingsEffect = settingsService.getSettings.pipe(
    Effect.map((settings) => settings.imageBank),
    Effect.mapError(() => imageBankError("configure", "Could not read image-bank settings.")),
  );

  const executeJson = Effect.fnUntraced(function* (
    request: HttpClientRequest.HttpClientRequest,
    operation: ImageBankError["operation"],
  ) {
    const response = yield* httpClient
      .execute(request)
      .pipe(Effect.mapError(() => imageBankError(operation, "Could not reach Microsoft.")));
    const body = yield* response.json.pipe(
      Effect.mapError(() =>
        imageBankError(operation, "Microsoft returned an invalid JSON response."),
      ),
    );
    return { status: response.status, body };
  });

  const executeBinary = Effect.fnUntraced(function* (url: string, token: string) {
    const response = yield* httpClient
      .execute(
        HttpClientRequest.get(url).pipe(
          HttpClientRequest.setHeader("authorization", `Bearer ${token}`),
        ),
      )
      .pipe(
        Effect.mapError(() => imageBankError("import", "Could not download the selected image.")),
      );
    if (response.status < 200 || response.status >= 300) {
      return yield* imageBankError(
        "import",
        `Microsoft Graph returned ${response.status} while downloading the image.`,
      );
    }
    const buffer = yield* response.arrayBuffer.pipe(
      Effect.mapError(() => imageBankError("import", "Could not read the downloaded image.")),
    );
    const bytes = new Uint8Array(buffer);
    if (bytes.length > MAX_IMAGE_BYTES) {
      return yield* imageBankError("import", "IMAGE_TOO_LARGE");
    }
    return {
      bytes,
      mimeType: response.headers["content-type"]?.split(";")[0] ?? "application/octet-stream",
    };
  });

  const graphJson = <A>(
    url: string,
    token: string,
    schema: Schema.Decoder<A>,
    operation: ImageBankError["operation"],
  ): Effect.Effect<A, ImageBankError> =>
    Effect.gen(function* () {
      const parsedUrl = yield* Effect.try({
        try: () => new URL(url),
        catch: () => imageBankError(operation, "Microsoft Graph returned an invalid URL."),
      });
      if (parsedUrl.origin !== "https://graph.microsoft.com") {
        return yield* imageBankError(operation, "Microsoft Graph returned an unexpected URL.");
      }
      const response = yield* executeJson(
        HttpClientRequest.get(parsedUrl.toString()).pipe(
          HttpClientRequest.setHeader("authorization", `Bearer ${token}`),
          HttpClientRequest.setHeader("accept", "application/json"),
        ),
        operation,
      );
      if (response.status < 200 || response.status >= 300) {
        const graphError = yield* Schema.decodeUnknownEffect(GraphErrorResponse)(
          response.body,
        ).pipe(Effect.orElseSucceed(() => null));
        return yield* imageBankError(
          operation,
          graphError?.error?.message ?? `Microsoft Graph returned ${response.status}.`,
        );
      }
      return yield* Schema.decodeEffect(schema)(response.body).pipe(
        Effect.mapError(() =>
          imageBankError(operation, "Microsoft Graph returned an unexpected response."),
        ),
      );
    }).pipe(Effect.scoped);

  const oauthJson = (url: string, params: Record<string, string>) =>
    executeJson(
      HttpClientRequest.post(url).pipe(HttpClientRequest.bodyUrlParams(params)),
      "sign-in",
    );

  const decodeTokenResponse = Effect.fnUntraced(function* (
    response: { readonly status: number; readonly body: unknown },
    failureMessage: string,
  ) {
    if (response.status < 200 || response.status >= 300) {
      const oauthError = yield* Schema.decodeUnknownEffect(OAuthErrorResponse)(response.body).pipe(
        Effect.orElseSucceed(() => null),
      );
      return yield* imageBankError(
        "sign-in",
        oauthError?.error_description ?? oauthError?.error ?? failureMessage,
      );
    }
    return yield* Schema.decodeUnknownEffect(OAuthTokenResponse)(response.body).pipe(
      Effect.mapError(() => imageBankError("sign-in", "Microsoft did not return sign-in tokens.")),
    );
  });

  const readTokenBundle = Effect.gen(function* () {
    const settings = yield* settingsEffect;
    const stored = yield* secretStore
      .get(TOKEN_SECRET)
      .pipe(
        Effect.mapError(() =>
          imageBankError("sign-in", "Could not read the saved Microsoft sign-in."),
        ),
      );
    if (Option.isNone(stored)) return null;
    const decoded = yield* Schema.decodeEffect(TokenBundleJson)(
      new TextDecoder().decode(stored.value),
    ).pipe(
      Effect.mapError(() =>
        imageBankError("sign-in", "The saved Microsoft sign-in is invalid. Sign in again."),
      ),
    );
    return decoded.tenantId === (settings.tenantId || "organizations") &&
      decoded.clientId === settings.clientId
      ? decoded
      : null;
  });

  const saveTokenBundle = Effect.fnUntraced(function* (bundle: TokenBundle) {
    const encoded = yield* Schema.encodeEffect(TokenBundleJson)(bundle).pipe(
      Effect.mapError(() =>
        imageBankError("sign-in", "Could not encode the Microsoft refresh token."),
      ),
    );
    yield* secretStore
      .set(TOKEN_SECRET, new TextEncoder().encode(encoded))
      .pipe(
        Effect.mapError(() =>
          imageBankError("sign-in", "Could not save the Microsoft refresh token."),
        ),
      );
  });

  const getAccessToken = Effect.fn("ImageBankService.getAccessToken")(function* () {
    const settings = yield* settingsEffect;
    const bundle = yield* readTokenBundle;
    if (!bundle)
      return yield* imageBankError("sign-in", "Sign in to the image bank in Settings first.");
    const now = yield* Clock.currentTimeMillis;
    const cached = yield* Ref.get(accessToken);
    if (cached && cached.expiresAt > now + 60_000) return cached.token;
    const response = yield* oauthJson(
      `https://login.microsoftonline.com/${encodeURIComponent(settings.tenantId || "organizations")}/oauth2/v2.0/token`,
      {
        client_id: settings.clientId,
        grant_type: "refresh_token",
        refresh_token: bundle.refreshToken,
        scope: TOKEN_SCOPES,
      },
    );
    const refreshed = yield* decodeTokenResponse(
      response,
      "Microsoft sign-in expired. Sign in again.",
    );
    const expiresAt = now + (refreshed.expires_in ?? 3600) * 1000;
    yield* Ref.set(accessToken, { token: refreshed.access_token, expiresAt });
    if (refreshed.refresh_token) {
      yield* saveTokenBundle({ ...bundle, refreshToken: refreshed.refresh_token });
    }
    return refreshed.access_token;
  });

  const getStatus: ImageBankService["Service"]["getStatus"] = Effect.gen(function* () {
    const settings = yield* settingsEffect;
    const bundle = yield* readTokenBundle;
    return {
      configured: settings.clientId.trim().length > 0 && settings.shareUrl.trim().length > 0,
      connected: bundle !== null,
      signedInAs: bundle?.signedInAs ?? null,
    };
  });

  const startLogin: ImageBankService["Service"]["startLogin"] = Effect.gen(function* () {
    const settings = yield* settingsEffect;
    if (!settings.clientId.trim()) {
      return yield* imageBankError(
        "configure",
        "Enter the Microsoft public-client application ID in Settings.",
      );
    }
    const tenantId = settings.tenantId || "organizations";
    const response = yield* oauthJson(
      `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/devicecode`,
      { client_id: settings.clientId, scope: TOKEN_SCOPES },
    );
    if (response.status < 200 || response.status >= 300) {
      const oauthError = yield* Schema.decodeUnknownEffect(OAuthErrorResponse)(response.body).pipe(
        Effect.orElseSucceed(() => null),
      );
      return yield* imageBankError(
        "sign-in",
        oauthError?.error_description ??
          oauthError?.error ??
          "Microsoft device sign-in could not start.",
      );
    }
    const authorization = yield* Schema.decodeUnknownEffect(DeviceCodeResponse)(response.body).pipe(
      Effect.mapError(() =>
        imageBankError("sign-in", "Microsoft did not return a device sign-in code."),
      ),
    );
    const intervalSeconds = Math.max(1, Math.trunc(authorization.interval ?? 5));
    const now = yield* Clock.currentTimeMillis;
    const expiresAt = now + authorization.expires_in * 1000;
    yield* Ref.set(pendingLogin, {
      deviceCode: authorization.device_code,
      tenantId,
      clientId: settings.clientId,
      expiresAt,
      intervalSeconds,
    });
    return {
      verificationUri: authorization.verification_uri,
      userCode: authorization.user_code,
      message: authorization.message ?? "Open the Microsoft sign-in page and enter this code.",
      expiresAt,
      intervalSeconds,
    };
  }).pipe(Effect.scoped);

  const pollLogin: ImageBankService["Service"]["pollLogin"] = Effect.gen(function* () {
    const pending = yield* Ref.get(pendingLogin);
    const now = yield* Clock.currentTimeMillis;
    if (!pending || now >= pending.expiresAt) {
      yield* Ref.set(pendingLogin, null);
      return { status: "expired" } as const;
    }
    const response = yield* oauthJson(
      `https://login.microsoftonline.com/${encodeURIComponent(pending.tenantId)}/oauth2/v2.0/token`,
      {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: pending.clientId,
        device_code: pending.deviceCode,
      },
    );
    if (response.status < 200 || response.status >= 300) {
      const failure = yield* Schema.decodeUnknownEffect(OAuthErrorResponse)(response.body).pipe(
        Effect.mapError(() =>
          imageBankError("sign-in", "Microsoft returned an invalid sign-in response."),
        ),
      );
      if (failure.error === "authorization_pending") return { status: "pending" } as const;
      if (failure.error === "slow_down") {
        const next = yield* Ref.modify(pendingLogin, (current) => {
          if (!current) return [null, null] as const;
          const updated = { ...current, intervalSeconds: current.intervalSeconds + 5 };
          return [updated, updated] as const;
        });
        return {
          status: "pending",
          ...(next ? { intervalSeconds: next.intervalSeconds } : {}),
        } as const;
      }
      if (failure.error === "expired_token") {
        yield* Ref.set(pendingLogin, null);
        return { status: "expired" } as const;
      }
      return yield* imageBankError(
        "sign-in",
        failure.error_description ?? "Microsoft sign-in was declined.",
      );
    }
    const tokens = yield* Schema.decodeUnknownEffect(OAuthTokenResponse)(response.body).pipe(
      Effect.mapError(() =>
        imageBankError("sign-in", "Microsoft did not return the sign-in tokens."),
      ),
    );
    if (!tokens.refresh_token)
      return yield* imageBankError("sign-in", "Microsoft did not return a refresh token.");

    const profileResult = yield* Effect.result(
      graphJson(
        `${GRAPH_ROOT}/me?$select=displayName,mail,userPrincipalName`,
        tokens.access_token,
        GraphProfile,
        "sign-in",
      ),
    );
    const profile = Result.isSuccess(profileResult) ? profileResult.success : null;
    const signedInAs =
      profile?.mail ?? profile?.userPrincipalName ?? profile?.displayName ?? "Microsoft account";
    const settings = yield* settingsEffect;
    yield* saveTokenBundle({
      tenantId: pending.tenantId,
      clientId: pending.clientId,
      refreshToken: tokens.refresh_token,
      signedInAs,
    });
    const expiresAt = now + (tokens.expires_in ?? 3600) * 1000;
    yield* Ref.set(accessToken, { token: tokens.access_token, expiresAt });
    yield* Ref.set(pendingLogin, null);
    yield* Ref.set(searchCache, null);
    return { status: "connected", signedInAs } as const;
  }).pipe(Effect.scoped);

  const resolveSharedRoot = Effect.fn("ImageBankService.resolveSharedRoot")(function* (
    token: string,
    shareUrl: string,
    operation: ImageBankError["operation"],
  ) {
    const parsedUrl = yield* Effect.try({
      try: () => new URL(shareUrl),
      catch: () => imageBankError(operation, "Enter a valid SharePoint or OneDrive sharing link."),
    });
    if (parsedUrl.protocol !== "https:") {
      return yield* imageBankError(operation, "The image-bank link must use HTTPS.");
    }
    const item = yield* graphJson(
      `${GRAPH_ROOT}/shares/${shareId(shareUrl)}/driveItem?$expand=listItem($expand=fields)`,
      token,
      GraphDriveItem,
      operation,
    );
    const driveId = item.parentReference?.driveId;
    if (!driveId || !item.id) {
      return yield* imageBankError(
        operation,
        "The sharing link did not resolve to a SharePoint image library.",
      );
    }
    if (!item.folder) {
      return yield* imageBankError(
        operation,
        "The SharePoint link must point to a library folder, not an individual image.",
      );
    }
    return { item, driveId };
  });

  const sampleColumns = Effect.fn("ImageBankService.sampleColumns")(function* (
    token: string,
    driveId: string,
    root: GraphDriveItem,
  ) {
    const pending = [root];
    const columns = new Map<string, string>();
    let visited = 0;
    while (pending.length > 0 && visited < 500 && columns.size < 100) {
      const folder = pending.shift();
      if (!folder?.id) continue;
      const page = yield* graphJson(
        `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}/children?$expand=listItem($expand=fields)&$top=100`,
        token,
        GraphResponse,
        "test",
      );
      for (const child of page.value ?? []) {
        visited += 1;
        if (child.folder) pending.push(child);
        for (const name of Object.keys(child.listItem?.fields ?? {})) {
          if (name) columns.set(name, displayFieldName(name) || name);
        }
      }
    }
    return [...columns].map(([name, displayName]): ImageBankColumn => ({ name, displayName }));
  });

  const testConnection: ImageBankService["Service"]["testConnection"] = Effect.gen(function* () {
    const settings = yield* settingsEffect;
    if (!settings.shareUrl.trim()) {
      return yield* imageBankError(
        "configure",
        "Enter the image-library sharing link in Settings.",
      );
    }
    const token = yield* getAccessToken();
    const { item, driveId } = yield* resolveSharedRoot(token, settings.shareUrl.trim(), "test");
    return {
      libraryName: item.name ?? "SharePoint image library",
      columns: yield* sampleColumns(token, driveId, item),
    };
  }).pipe(Effect.scoped);

  const loadItems = Effect.fn("ImageBankService.loadItems")(function* (
    token: string,
    settings: ImageBankSettings,
  ) {
    const { item: root, driveId } = yield* resolveSharedRoot(
      token,
      settings.shareUrl.trim(),
      "search",
    );
    const folders: GraphDriveItem[] = [root];
    const items: CachedItem[] = [];
    let visited = 0;
    while (folders.length > 0 && visited < MAX_LIBRARY_ITEMS) {
      const folder = folders.shift();
      if (!folder?.id) continue;
      let nextUrl: string | null =
        `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}/children?$expand=listItem($expand=fields)&$top=200`;
      while (nextUrl && visited < MAX_LIBRARY_ITEMS) {
        const page: GraphResponse = yield* graphJson(nextUrl, token, GraphResponse, "search");
        for (const item of page.value ?? []) {
          visited += 1;
          if (item.folder) folders.push(item);
          else {
            const candidate = toCandidate(item, settings.fieldMapping);
            if (candidate && item.id) items.push({ candidate, driveId, itemId: item.id });
          }
          if (visited >= MAX_LIBRARY_ITEMS) break;
        }
        nextUrl = page["@odata.nextLink"] ?? null;
      }
    }
    return items;
  });

  const itemsForSearch = Effect.fn("ImageBankService.itemsForSearch")(function* (
    token: string,
    settings: ImageBankSettings,
  ) {
    const now = yield* Clock.currentTimeMillis;
    const key = `${settings.shareUrl}\u0000${Object.values(settings.fieldMapping).join("\u0000")}`;
    const cached = yield* Ref.get(searchCache);
    if (cached && cached.key === key && cached.expiresAt > now) return cached.items;
    const items = yield* loadItems(token, settings);
    yield* Ref.set(searchCache, { expiresAt: now + CACHE_TTL_MS, key, items });
    return items;
  });

  const search: ImageBankService["Service"]["search"] = (input) =>
    Effect.gen(function* () {
      const settings = yield* settingsEffect;
      const token = yield* getAccessToken();
      const items = yield* itemsForSearch(token, settings);
      const terms = normalize(input.query).split(/\s+/).filter(Boolean);
      const themeFilter = input.themes ? normalize(input.themes) : "";
      const workfieldFilter = input.workfields ? normalize(input.workfields) : "";
      const typeFilter = input.imageType ? normalize(input.imageType) : "";
      const ranked = items.flatMap(({ candidate }) => {
        const themeText = normalize(candidate.themes.join(" "));
        const workfieldText = normalize(candidate.workfields.join(" "));
        const typeText = normalize(candidate.imageType.join(" "));
        if (themeFilter && !themeText.includes(themeFilter)) return [];
        if (workfieldFilter && !workfieldText.includes(workfieldFilter)) return [];
        if (typeFilter && !typeText.includes(typeFilter)) return [];
        const content = normalize(
          [
            candidate.name,
            candidate.fileName,
            ...candidate.themes,
            ...candidate.imageType,
            ...candidate.workfields,
            ...candidate.keywords,
          ].join(" "),
        );
        const score = terms.reduce((sum, term) => sum + (content.includes(term) ? 1 : 0), 0);
        return score > 0 ? [{ candidate, score }] : [];
      });
      ranked.sort(
        (left, right) =>
          right.score - left.score || left.candidate.name.localeCompare(right.candidate.name),
      );
      return { candidates: ranked.slice(0, input.limit ?? 8).map(({ candidate }) => candidate) };
    }).pipe(Effect.scoped);

  const importAsset: ImageBankService["Service"]["import"] = (input: ImageBankImportInput) =>
    Effect.gen(function* () {
      const settings = yield* settingsEffect;
      const token = yield* getAccessToken();
      const items = yield* itemsForSearch(token, settings);
      const item = items.find((entry) => entry.candidate.assetId === input.assetId);
      if (!item)
        return yield* imageBankError(
          "import",
          "That image was not found in the connected library. Search again.",
        );
      if (item.candidate.restricted && input.allowRestricted !== true) {
        return yield* imageBankError(
          "import",
          "This image has a usage restriction. Confirm it is permitted before importing with allowRestricted=true.",
        );
      }
      const baseUrl = `${GRAPH_ROOT}/drives/${encodeURIComponent(item.driveId)}/items/${encodeURIComponent(item.itemId)}`;
      let downloaded: { readonly bytes: Uint8Array; readonly mimeType: string } | null = null;
      let fromThumbnail = item.candidate.sizeBytes > MAX_IMAGE_BYTES;
      if (!fromThumbnail) {
        const original = yield* Effect.result(executeBinary(`${baseUrl}/content`, token));
        if (Result.isSuccess(original)) downloaded = original.success;
        else if (original.failure.detail === "IMAGE_TOO_LARGE") fromThumbnail = true;
        else return yield* original.failure;
      }
      if (fromThumbnail) {
        downloaded = yield* executeBinary(`${baseUrl}/thumbnails/0/large/content`, token).pipe(
          Effect.mapError(() =>
            imageBankError(
              "import",
              "The original image is too large and Microsoft did not provide a usable thumbnail.",
            ),
          ),
        );
      }
      if (!downloaded)
        return yield* imageBankError("import", "The selected image could not be downloaded.");
      const fileName = fromThumbnail
        ? `${item.candidate.fileName.replace(/\.[^.]+$/, "")}-thumbnail.jpg`
        : safeFileName(item.candidate.fileName);
      const itemSuffix = item.itemId.replace(/[^a-zA-Z0-9_-]/g, "").slice(-12) || "image";
      const relativePath =
        input.relativePath?.trim() || `assets/image-bank/${itemSuffix}-${fileName}`;
      if (
        relativePath.startsWith("/") ||
        relativePath.startsWith("\\") ||
        /^[a-zA-Z]:/.test(relativePath) ||
        relativePath.split(/[\\/]/).some((part) => part === ".." || part === ".")
      ) {
        return yield* imageBankError(
          "import",
          "The import path must stay inside the project workspace.",
        );
      }
      const asset: ImageBankImportedAsset = {
        relativePath,
        fileName,
        mimeType: downloaded.mimeType.startsWith("image/")
          ? downloaded.mimeType
          : item.candidate.mimeType,
        sourceItemId: item.itemId,
        sourceUrl: item.candidate.webUrl,
        themes: item.candidate.themes,
        imageType: item.candidate.imageType,
        workfields: item.candidate.workfields,
        keywords: item.candidate.keywords,
        restrictions: item.candidate.restrictions,
      };
      return { asset, contentsBase64: Buffer.from(downloaded.bytes).toString("base64") };
    }).pipe(Effect.scoped);

  const disconnect: ImageBankService["Service"]["disconnect"] = Effect.gen(function* () {
    yield* secretStore
      .remove(TOKEN_SECRET)
      .pipe(
        Effect.mapError(() =>
          imageBankError("disconnect", "Could not remove the saved Microsoft sign-in."),
        ),
      );
    yield* Ref.set(accessToken, null);
    yield* Ref.set(pendingLogin, null);
    yield* Ref.set(searchCache, null);
    return yield* getStatus;
  });

  return ImageBankService.of({
    getStatus,
    startLogin,
    pollLogin,
    testConnection,
    disconnect,
    search,
    import: importAsset,
  });
});

export const layer = Layer.effect(ImageBankService, make);
