# Release checklist

> For maintainers. For installation and updates, see [the user guide](../user/).

## Release workflow

The release workflow is `.github/workflows/release.yml`. It builds the web client and server once,
then builds five CLI archives through `.github/workflows/release-cli.yml`:

- macOS arm64
- Linux x64 and arm64
- Windows x64 and arm64

The GitHub Release contains those archives and `SHA256SUMS`. The workflow publishes the same CLI
builds to npm as `t3` and the platform packages under `@t3code/`. It deploys the hosted web app on
stable and nightly releases, and deploys the marketing site on nightlies.

The workflow supports three channels:

- **Stable** runs from a version tag or a manual dispatch. A manual release builds the commit from
  the latest published nightly unless `version` is supplied.
- **Nightly** runs every 30 minutes when there are new commits and at least six hours have passed
  since the previous nightly. A manual dispatch skips those checks.
- **Preview** is a manual prerelease for exercising the publishing flow. It publishes public GitHub
  and npm prereleases, but does not deploy either website or move stable aliases.

All dispatch channels publish real release artifacts. There is no non-publishing release run.
Stable tags with a prerelease suffix are GitHub prereleases. Only plain `X.Y.Z` releases become the
latest GitHub release. Nightly and preview releases never become latest.

Stable releases commit aligned package versions back to `main`. Nightlies and previews do not.
Automatically generated notes compare against the previous release in the same channel.

## Release requirements

Configure npm trusted publishing for `t3` and each `@t3code/t3-<platform>-<arch>` package. Use
GitHub Actions, this repository, and `.github/workflows/release.yml` as the trusted publisher. The
workflow dry-runs all package publishes before making any package live.

Required GitHub Actions secrets and variables:

- `RELEASE_APP_ID` and `RELEASE_APP_PRIVATE_KEY` for stable version updates on `main`.
- `VERCEL_TOKEN`, `VERCEL_ORG_ID`, and `VERCEL_PROJECT_ID` for the hosted web deployment.
- `VERCEL_TOKEN` and `VERCEL_ORG_ID` for the marketing deployment. The workflow looks up the
  `t3code-marketing` project; `VERCEL_TEAM_SLUG` is optional.
- The `production` environment's Clerk and relay values, read by the
  [T3 Connect release configuration](./connect-setup.md).

Code signing is optional. The CLI workflow signs macOS archives when `CSC_LINK` and
`CSC_KEY_PASSWORD` are configured, and signs Windows executables when the Azure Trusted Signing
secrets are present. Unsigned archives are still published when those credentials are absent.

The Vercel web project root must be `apps/web`; the marketing project root must be
`apps/marketing`. Automatic Git deployments are disabled in each project's `vercel.ts` file.

## Hosted web deployment

The release workflow deploys the hosted web app after publishing the GitHub Release. Stable releases
alias the deployment to `latest.app.t3.codes` and `app.t3.codes`. Nightlies update
`nightly.app.t3.codes`. The router at `app.t3.codes` stores a channel cookie and sends later requests
to the selected alias.

Required Vercel domains:

- `app.t3.codes`
- `latest.app.t3.codes`
- `nightly.app.t3.codes`

The optional variables `T3CODE_WEB_ROUTER_URL`, `T3CODE_WEB_LATEST_DOMAIN`, and
`T3CODE_WEB_NIGHTLY_DOMAIN` override the defaults. The Vercel web project root must stay
`apps/web`, and its `vercel.ts` routing rules must be deployed before enabling the router domain.

The marketing project deploys only after a nightly release. Stable releases can promote an older
nightly commit, so they do not deploy the marketing site.

## Server update ordering

Connected servers update to the exact version requested by the web client. The matching `t3` npm
version must be available before that client reaches users. Keep the workflow dependencies in this
order:

1. `publish_cli` publishes the version to npm.
2. `release` publishes the GitHub Release.
3. `deploy_web` moves the hosted web channel to the new version.

For a release smoke check, confirm `npm view t3@<version> version` returns the expected version.
Connect the new web client to a server on the previous version and confirm its server update reaches
the matching release. If the release adds database migrations, verify the update applies them and
reconnects. Restore the database snapshot if the update trial fails.

## T3 Connect relay deployment

The relay is versioned and deployed separately from web and CLI releases. Stable and nightly builds
use the production relay so linked environments remain available when users change channels.

`.github/workflows/deploy-relay.yml` deploys Alchemy stage `prod` on pushes to `main`. The
`production` GitHub Actions environment supplies the Clerk public configuration and relay URL to
the release build. See [T3 Connect setup](./connect-setup.md) and the
[relay deployment guide](../../infra/relay/README.md#deployment) for configuration.

Developers deploy personal relay stages locally rather than through pull request automation:

```sh
vp run --filter t3code-relay deploy -- --stage "$USER" --env-file .env.local
```

### Managed tunnel cleanup rollout

Keep `RELAY_TUNNEL_CLEANUP_MODE=off` for the first production deploy. That deploy applies the
allocation migration and adds the recovery endpoints. Release a server version that registers
recovery before enabling cleanup, because older hosts cannot replace a deleted tunnel after wake.

1. Deploy the relay and migration with cleanup `off`.
2. Release the server and confirm current hosts register recovery.
3. Set `dry-run`, force a relay deploy, and review the sweep counters in Axiom over several sweeps.
4. Run the disposable-host canary below.
5. Set `enabled` only after the canary recovers without a server restart.

The job runs every five minutes with a five-minute grace period for disconnected tunnels. A tunnel
is usually removed five to ten minutes after it goes down. Tunnels that never connected wait an
hour. A sweep attempts at most 100 deletions, so a backlog takes longer. Changing the cleanup mode
requires a forced relay deploy.

To roll back, set cleanup to `off` and force a relay deploy before downgrading a host. Keep the
recovery endpoints deployed while released server builds use them. The nullable allocation columns
can remain.

### Disposable-host canary

Run this against a disposable relay stage, Cloudflare account, host, and T3 home. Keep production
cleanup at `off` or `dry-run` until the canary passes.

1. Deploy the disposable stage with cleanup `dry-run`. Link an environment from the web app and
   confirm its tunnel is healthy and recovery is registered.
2. Stop the host, then restart its T3 home on a different local port. Confirm the public hostname
   reaches the new port.
3. Link a second environment with a server that predates recovery registration. Stop its managed
   `cloudflared` process and confirm dry-run counts it as `skippedLegacy`.
4. Stop the first environment's `cloudflared` process and confirm dry-run counts its tunnel under
   `wouldDelete`.
5. Enable cleanup on the disposable stage and force a relay deploy. Confirm the first tunnel is
   deleted while the legacy tunnel remains.
6. Resume the first process. Confirm the running server replaces its tunnel and becomes reachable
   at the same hostname without a restart.
7. Resume the legacy process and confirm it reconnects.
8. Repeat with a physical sleep and wake cycle on a disposable laptop before broad rollout.

Confirm each PID belongs to the disposable host before stopping or resuming it. Do not stop a
daily-use T3 server.
