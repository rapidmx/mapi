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
- **Commit discipline.** Don't `git commit` unless explicitly asked for *that specific piece of
  work*. An autonomous-execution/"commit as you go" approval given for one approved plan (e.g. via
  plan mode) is scoped to that plan only — it does not carry forward to later, separate requests in
  the same session, even ones that look similar in kind (a follow-up review-and-fix pass, a
  refactor, a new feature), and even after a full review-and-fix cycle with passing tests. Default
  to leaving changes staged/unstaged and saying so; only commit automatically within the exact
  scope of a plan that was explicitly approved as autonomous. If unsure whether new work falls
  inside that scope, treat it as outside and ask.
- **Commit message style: a flat list of one-line, verb-led items — no summary/title line, no
  `-`/`*` bullet markers.** This isn't just a style preference — it's dictated by how `release`
  (`@rapidrest/cli`) actually builds `CHANGELOG.md`. `collectChangelogBullets`/
  `classifyChangelogLine` (that repo's `src/lib/release.ts`) parse `git log --pretty=format:%B` and
  treat **every non-blank line of a commit's full message as its own changelog bullet** — there is
  no subject/body distinction. A conventional "short imperative subject + blank line + prose body"
  commit therefore leaks one changelog bullet per body sentence, and a `-`/`*`-prefixed line breaks
  `classifyChangelogLine`'s verb detection (it reads the line's first whitespace-delimited word as
  the verb; a leading `-` defeats that lookup and the dash leaks into the changelog text as
  `"- - Added foo"`). Correct format:
  - No separate summary/title line — if a commit needs an overview, that overview is itself just
    one more flat line, not a heading distinct from the rest.
  - No bullet-marker prefix of any kind — write bare lines.
  - Lead each line with an imperative verb where it fits: `Add`/`Fix`/`Remove` (and `-ing` forms)
    are recognized and become `Added`/`Fixed`/`Removed` entries; `Configuring`/`Converting`/
    `Refactoring`/`Updating`/etc. become `Changed`. Anything else still works, defaulting to
    `Changed` verbatim — see `CHANGELOG_VERB_REWRITES` in that repo's `src/lib/release.ts` for the
    full map.
  - A blank line before a trailing git trailer (`Co-Authored-By:`, `Signed-off-by:`, etc.) is fine
    — trailers matching `CHANGELOG_NOISE_PATTERNS` are dropped from the changelog — but nothing
    else should follow the item list.
  This mirrors JP's standing convention across his other repos; copy this exact rule verbatim into
  each sibling repo's own NOTES.md rather than paraphrasing it, since the paraphrase is what caused
  this to be gotten wrong in the first place (see `@rapidrest/cli`'s own NOTES.md, 2026-09-07 entry,
  for the full incident writeup and the `CHANGELOG_NOISE_PATTERNS` fix that accompanied it).

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
