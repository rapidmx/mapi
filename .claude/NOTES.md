# mapi — Design Decisions & Session Notes

This file exists so that Claude sessions working in this repo don't re-litigate settled
decisions or re-discover the same issues from scratch. It is local to this repo (not tied to
any one machine's global Claude memory), so it travels with the code.

**Maintenance rule:** when a standing decision changes, update the section below in place
(don't just append a contradiction lower down). When a new investigation/session produces a
decision, finding, or reverted approach worth remembering, add a dated entry under Session Log.
Keep entries terse — this is a reference, not a transcript.

## Standing design decisions & constraints

- **Vulnerability/review threat model: externally-exploitable only.** Only count issues reachable
  from a downstream, untrusted HTTP client hitting a service built on this package (anonymous or
  low-privilege caller). Do NOT flag developer-only footguns or purely theoretical races with no
  concrete external trigger path.
- **Commit discipline.** Don't `git commit` unless explicitly asked, even after a full
  review-and-fix cycle with passing tests. Leave changes staged/unstaged and say so.
- **Commit message style: concise, one line per task/bug/feature — no verbose prose.** A commit
  message is a short list of one-line bullets, one per item. This mirrors JP's standing convention
  across his other repos.

## Session Log

### 2026-09-06 — Repo split: `@rapidrest/mail` → four RapidMX packages

- **This repo is `@rapidmx/mapi`**, carved out of the former monolith `@rapidrest/mail`
  (`d:\github\rapidrest\mail`, still present there for reference/history) — was `src/mapi`/
  `test/mapi` there, moved to this repo's own root (not nested under a `mapi/` folder).
- Depends on [`@rapidmx/restapi`](https://github.com/RapidMX/restapi) (real published `^0.1.0`,
  not a local `portal:` link — see below) for the mailbox/folder/message/contact/calendar models,
  `resolveCallerMailboxUid`/`RecoverableRepoUtils`/`sendComposedMime` REST-layer helpers,
  `BlobStore`, and the scan pipeline.
- **Built from the start with two lessons already learned from the `activesync`/`autodiscover`
  splits (see their own NOTES.md for the full diagnosis of each) — do not regress either**:
  1. `vitest.config.ts`'s `ssr.noExternal` lists `@rapidmx/restapi` alongside
     `@rapidrest/service-core`/`@rapidrest/core`. Without it, Vite's SSR pipeline loads a second,
     natively-required copy of the framework packages for anything reached *through*
     `@rapidmx/restapi`, breaking `ModelUtils`'s static SQL-typeorm state (and, more generally, any
     `instanceof`-based framework behavior) between the two copies - manifesting as bare 500s from
     otherwise-correct code.
  2. `test/server-{mongo,sql}/models/index.ts` uses a **named** re-export of just the 5 model
     classes MAPI needs (`Folder`/`Message`/`Mailbox`/`Contact`/`CalendarEvent`), not a wildcard
     `export * from "@rapidmx/restapi/mongo"` - the wildcard form also pulls in every REST
     route/job class restapi bundles alongside its models, which the ClassLoader would then
     discover and try to initialize with no configured dependencies for them.
- Confirmed via a grep of `test/routes/{mongo,sql}/MapiEmsmdbRoute.test.ts` that (unlike
  `activesync`, which needed `FolderRoute`/`MessageRoute` REST mounts for its "renamed/deleted via
  REST" test scenarios) none of MAPI's own tests call out to a REST endpoint outside its own
  `/mapi/emsmdb`/`/mapi/nspi` base paths - no extra route mount files were needed here.
- For the original design rationale behind the ROP codec, the named-property registry, the
  calendar recurrence/timezone/GlobalObjectId codecs, and every other decision baked into this
  code, see the monolith's own `.claude/NOTES.md` (`d:\github\rapidrest\mail`) - that history
  wasn't duplicated here since it predates this repo's existence.
