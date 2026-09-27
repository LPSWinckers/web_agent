# Remote access

Connect a browser to T3 Code running on another machine. The host must stay on and reachable while
you work.

## T3 Connect

T3 Connect makes a server reachable without router port forwarding. On the host, run:

```bash
t3 connect
```

Follow the sign-in prompts. Setup offers a [background service](./background-service.md). If you
skip it, start the server with `t3 serve`. Signing in alone does not make the host reachable.

Open [app.t3.codes](https://app.t3.codes) in a browser signed in to the same T3 Connect account,
then choose the environment. For an SSH host, the CLI prints a browser link and a short code. Open
the link, confirm the code, and approve the connection. The CLI continues without forwarding an
OAuth callback port.

T3 Connect renews access credentials when needed without disconnecting a healthy connection. Pull
request diffs and provider settings keep working after the previous credential expires. A failed
renewal affects that request; it does not disconnect an otherwise healthy conversation.

## Pair over a LAN or private network

Use direct pairing when the browser can reach the host's network address. Start the server with:

```bash
t3 serve --host <private-ip>
```

If a server is already running, create a fresh link without restarting it:

```bash
t3 pair
```

Open the pairing URL in a browser and add the environment under **Settings → Connections**. A
loopback address such as `127.0.0.1` reaches only the device opening the link.

Pairing authorizes a browser for future connections. Create a fresh one-time link for each new
browser; you do not need the original token to reconnect. Links created in Settings can only be
copied from the browser that created them while its Connections page stays open. If you leave or
reload that page, create another link to share.

### Balance new threads across machines

Auto balance is off by default. Enable it in **Settings → Connections → Load balancing** to choose a
machine automatically for new threads in projects grouped across connected environments. The
section appears once two or more machines are switched on.

Each machine starts at **Normal**. Choose **Prefer** to favor it when it has CPU and memory
available, **Less often** to reduce its share, or **Manual only** to exclude it from automatic
selection. These are preferences, not fixed traffic percentages. Preferences are saved in the
browser.

The composer checks eligible machines when choosing a draft's environment, then keeps that choice
stable. Choose **Auto balance** again to check current resources, or choose a specific machine to
override it. Choosing a branch or worktree also keeps the draft on that machine. Existing threads
stay where they started. If resource checks are unavailable or all eligible machines are full,
choose a machine manually to continue.

### Tailscale HTTPS

Join the host and browser device to the same tailnet. Start the server with Tailscale HTTPS:

```bash
t3 serve --tailscale-serve
```

For a running server, create a pairing link with:

```bash
t3 pair --tailscale
```

The pairing link uses an address such as `https://machine.tailnet.ts.net/`. The mapping created by
`pair --tailscale` persists across restarts. Remove its default-port mapping with:

```bash
tailscale serve --https=443 off
```

If that port is already in use, choose another with `--tailscale-serve-port`. See `t3 pair --help`
for other pairing options.

## Hosted web app

[app.t3.codes](https://app.t3.codes) needs an HTTPS endpoint. It connects directly to your server; a
hosted pairing link does not make an unreachable backend reachable or convert HTTP to HTTPS.

For a plain HTTP LAN endpoint, open the direct pairing URL in a browser that can reach it. Use
`https://` in the URL when the server uses HTTPS.

## Manage or revoke access

On the host, **Settings → Connections** lets authorized administrators create pairing links and
revoke browser sessions. Revoking an unused link prevents new pairings; revoke a session to remove
its existing access. Command-line management is available through `t3 auth --help`.

A session with an open connection stays listed after its access credential expires.

To remove an environment from T3 Connect, open your account menu's **T3 Connect** page and choose
**Deregister**. This revokes its cloud access and frees its host space even when the environment is
offline or has been wiped.

When idle tunnel cleanup is enabled, T3 Connect removes a linked environment's tunnel after it stays
offline for several minutes. The environment stays linked and keeps the same address. When the host
starts again or wakes, T3 Connect creates a replacement tunnel on its own. You do not need to pair
again. Cleanup usually runs five to ten minutes after the tunnel goes down.

On a command-line host, `t3 connect unlink` disables exposure while retaining your login;
`t3 connect logout` also clears that login. Background-service [removal](./background-service.md#manage-the-service)
is separate.

Treat pairing URLs and authorization codes as passwords. Do not include them in screenshots, logs,
or bug reports.

## T3 Connect troubleshooting

Run `t3 connect status` on the host to inspect saved authorization and link configuration. It is not
a live reachability check. If the environment appears offline, run `t3 service status` and read the
displayed log. If it disappears when SSH closes, see
[background-service troubleshooting](./background-service.md#troubleshooting).

| Error                                                     | Recovery                                                                                                                                    |
| --------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| `environment_link_limit_exceeded` or managed tunnel limit | Deregister an unused environment, then restart T3 Code on the host.                                                                         |
| `auth_invalid` or `invalid_bearer`                        | Run `t3 connect login`. If credentials were revoked, run `t3 connect logout`, then `t3 connect` again. Restart the server after signing in. |
| Expired or invalid link proof                             | Check the host's date and time, update T3 Code, then restart it.                                                                            |
| HTTP 403 without a recognized error                       | Check relay access, proxies, and firewall rules. Keep any Cloudflare Ray ID for a bug report.                                               |
| HTTP 408, 429, or 5xx                                     | Check network and relay availability. Startup retries temporary failures for up to ten minutes.                                             |

After fixing a permanent rejection, restart the host's server. On Linux, use
`systemctl --user restart t3code.service` for the background service. For a foreground server, stop
it and run `t3 serve` again with your usual options. Include the diagnostic message and trace ID
when reporting a persistent failure.

For a connection that still fails after linking, check the date and time on both devices. For server
version warnings, follow [Updating T3 Code](./updating.md).
