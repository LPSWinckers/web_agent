import React from "react";
import ReactDOM from "react-dom/client";
import { createBrowserHistory } from "@tanstack/react-router";

import "./index.css";

import { hasCloudPublicConfig } from "./cloud/publicConfig";
import { getRouter } from "./router";
import { AppRoot } from "./AppRoot";
import { clearChunkReloadGuard, reloadOnceForChunkLoadError } from "./lib/chunkReloadGuard";

const router = getRouter(createBrowserHistory());

const clerkPublishableKey = import.meta.env.VITE_CLERK_PUBLISHABLE_KEY as string | undefined;

// A failed split-chunk fetch usually means the hashed assets went stale under
// a deploy; one guarded reload picks up the fresh index.html.
let chunkLoadFailed = false;
let reloadScheduled = false;
window.addEventListener("vite:preloadError", (event) => {
  chunkLoadFailed = true;
  if (reloadOnceForChunkLoadError()) {
    reloadScheduled = true;
    event.preventDefault();
  }
});

const app = <AppRoot router={router} />;

// Load Clerk only for hosted builds, keeping it out of the startup graph for
// local-mode users.
const managedAuthShellModule =
  clerkPublishableKey && hasCloudPublicConfig()
    ? import("./components/clerk/BrowserManagedAuthShell")
    : null;

// The index.html boot splash lives inside #root, and React's first commit
// clears it. Resolve everything that first commit needs, the selected
// managed-auth runtime and the initial route's split chunks, before
// rendering, so the splash holds until real UI paints instead of dropping to
// a blank window while chunks download.
export const startup = Promise.all([
  managedAuthShellModule?.then((module) => module.default) ?? null,
  router.load(),
])
  .then(([ManagedAuthShell]) => {
    // A route chunk failure still resolves router.load(): the error is parked in
    // the lazy component and surfaces through the route error boundary. Skip the
    // paint when a reload is on its way, and only re-arm the guard after a boot
    // that fetched every chunk it asked for.
    if (reloadScheduled) return;
    if (!chunkLoadFailed) clearChunkReloadGuard();
    ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
      <React.StrictMode>
        {ManagedAuthShell && clerkPublishableKey ? (
          <ManagedAuthShell publishableKey={clerkPublishableKey}>{app}</ManagedAuthShell>
        ) : (
          app
        )}
      </React.StrictMode>,
    );
  })
  .catch((error: unknown) => {
    // Let the bootstrap entry show the error unless a reload is already scheduled.
    if (reloadScheduled) return;
    throw error;
  });
