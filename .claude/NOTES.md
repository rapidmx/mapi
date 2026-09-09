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

### 2026-09-08 — Synced against `@rapidmx/restapi` 0.2.0→0.3.x: fixed a stale peer-dependency range and a
broken test mock, added Contacts/Tasks browsing + deferred-send/read-receipt Submit support

Prompted by a "review restapi's latest activity and implement missing/broken MAPI protocol features" request.
Reviewed every restapi commit since this repo's own initial split (EAS-sync fields, TaskList, mail filters,
scheduled send, signatures/OOF, distribution lists, transport rules, audit log, domains, focused inbox,
bookings, plus-addressing, branding, read/delivery receipts) and filtered to what's actually MAPI/HTTP-protocol
relevant for a real Outlook client - **deliberately excluded** Branding/Bookings/TransportRule/AuditLog/Domains/
DNS-setup/DistributionList-admin/MailFilterRule-as-Outlook-Rules and EAS-only fields, since none of those are
things a MAPI client consumes (they're REST-admin/webmail/EAS-only surface) - implementing Outlook "Rules" ROPs
against `MailFilterRule` in particular was considered and set aside as a separate, much larger effort (real
`RopGetRulesTable`/`RopModifyRules` support, still listed as a documented gap) rather than folded into this pass.

- **Found two genuine regressions from the version drift** (peer dep pinned `0.1.x` in `package.json` while
  `devDependencies`/tests already ran against `^0.3.0` - a real installability bug for any downstream consumer
  on a modern restapi): bumped the peer range to `^0.3.0`. Separately, `test/rop/RopSubmitMessageHandler.test.ts`'s
  own `makeScanPipeline()` mock predated restapi's new `deriveConversationId()` call inside `scanAndRelay()`
  (added for `Message.conversationId`) - the mock's `run()` result had no `references`/`inReplyTo` fields, so
  `references[0]` threw on `undefined`. Fixed by adding both fields to the mock (production code was never
  actually broken - a real `ScanPipeline.run()` always populates them - only this repo's own test double was
  stale).
- **Added Contacts/Tasks folder browsing** - the biggest actual protocol gap found, not called out in the
  README's own "documented gaps" list before this (an oversight, not a deliberate exclusion): `RopGetContentsTable`
  already special-cased `FolderType.CALENDAR` to route to `calendarEventRepo` instead of `messageRepo`, but
  `CONTACTS`/`TASKS` folders fell through to `resolveFolderMessages()` and would always show empty, since
  `Contact`/`Task` are separate restapi entities, never `Message` rows. New `ContactTarget.ts`/`TaskTarget.ts`
  mirror `CalendarEventTarget.ts`'s exact shape (`resolveXInfo`/`resolveFolderXs`, `"contact:"`/`"task:"` target
  prefixes reusing `MessageTarget.assignOrGetMid`'s MID registry as-is). Wired through `RopGetContentsTableHandler`,
  `RopOpenMessageHandler` (subject/title resolution), and `PropertyResolvers.resolvePropertyValues` (new
  `contactValueFor`/`taskValueFor`). Task's own MAPI properties (`PidLidTaskStatus`/`PercentComplete`/
  `TaskDueDate`/`TaskComplete`, `PSETID_Task`) are named properties per real `[MS-OXOTASK]` - new
  `TaskNamedProperties.ts` mirrors `CalendarNamedProperties.ts`'s pattern; verified every LID/GUID against
  Microsoft's own Learn docs before implementing (`PidLidTaskStatus` = `0x8101`, `PercentComplete` = `0x8102`,
  `TaskDueDate` = `0x8105`, `TaskComplete` = `0x811C`, `PSETID_Task` = `00062003-0000-0000-c000-000000000046`) -
  Contact's own fields are all plain fixed `PidTag`s (`GivenName`/`Surname`/`EmailAddress`/
  `BusinessTelephoneNumber`/`CompanyName`/`Title`), no named-property lookup needed.
  **Deliberately read-only**: no `RopSetProperties`/`RopSaveChangesMessage` write-back for either kind in this
  pass (mirrors this package's own existing "browse only" precedent, e.g. no folder rename/move) - the REST API
  remains the way to create/edit Contacts/Tasks.
  **`contactRepo`/`taskRepo`/`contactClass`/`taskClass` were added to `RopContext` as *optional* fields**,
  not required ones - making them required would have forced every one of this codebase's ~15 existing
  `RopContext`-object-literal test fixtures (predating this feature) to grow four new fields each for no
  behavioral reason. `BaseMapiEmsmdbRoute` always populates them for real; every consumer degrades a
  `"contact:"`/`"task:"` target to empty rows/type-appropriate defaults when the repo is absent from a test
  double, the same "missing data, not a crash" stance already used for a vanished row.
  Added a real end-to-end HTTP+DB integration test in both `test/routes/mongo/MapiEmsmdbRoute.test.ts` and its
  `sql/` counterpart (Logon → OpenFolder(root) → GetHierarchyTable → SetColumns([DisplayName, FolderId]) →
  QueryRows to discover the Contacts/Tasks FID - neither is in `RopLogon`'s fixed 13-folder `FolderIds` list,
  unlike Inbox - then OpenFolder → GetContentsTable → SetColumns → QueryRows against it), plus full unit
  coverage for the new target/resolver files and the extended dispatch/handler branches.
- **Added `PidTagDeferredSendTime`/`PidTagReadReceiptRequested` support to `RopSubmitMessage`** - two documented
  gaps a real Outlook client's own Send dialog exercises routinely ("Do not deliver before" / "Request a read
  receipt"), now tracked by `RopSetPropertiesHandler` and honored in `RopSubmitMessageHandler`: a future deferred
  time parks the composed-but-unrelayed message in Outbox with `Message.scheduledSendTime` set (mirroring
  `BaseMessageRoute.send()`'s own identical branch, letting `ScheduledSendJob` relay it later via the same
  `scanAndRelay()` call), and a read-receipt request attaches a real `Disposition-Notification-To` header via
  `MailComposer`'s own `headers` option (no need to reach into restapi's internal `MimeHeaderUtils`) plus records
  `requestReceipt` on the persisted message. Deliberately **not** wired to this mailbox's own
  `alwaysRequestReceiptInternal`/`External` defaults - those apply to the REST/webmail compose path only,
  consistent with this handler's existing "no `RopModifyRecipients`, no address-book resolution" pragmatic scope.
- **Considered and explicitly deferred**: exposing `Message.inferenceClassification` (Focused Inbox) via
  `PidTagInferenceClassification` - restapi's own doc comment says this field's values were deliberately chosen
  to match what MAPI exposes, an open invitation to wire it up - but repeated web searches for this property's
  real fixed tag ID/wire-value encoding came back inconclusive (no canonical `[MS-OXPROPS]`/`[MS-OXCMSG]` page
  found, unlike every other property this session touched, all of which were verified against a real Microsoft
  Learn page first). Implementing a guessed tag/value mapping for a real protocol field risked silently
  mis-rendering a genuine Outlook client's Focused/Other view - worse than not exposing it at all - so this was
  left as a documented gap in the README rather than shipped on a guess. Revisit if a reliable source for the
  real encoding turns up.
- Full suite: 447/447 passing (up from 412 pre-session's own count after fixing the 5 pre-existing failures,
  net +35 new tests), 100% statement/function/line coverage, 98.82% branch (comfortably above this package's own
  95% floor). `yarn build`/`yarn tsc --noEmit`/`yarn lint` all clean. Not committed - left staged/unstaged per
  the standing commit-discipline rule.

### 2026-09-09 — Two-agent adversarial review + fix pass: DoS, correctness, and O(n²) session-registry bugs

Prompted by "review this codebase using two adversarial agents, look for correctness, bugs, vulnerabilities and
bottlenecks" followed by "fix everything." Ran two independent `general-purpose` agents in parallel (one
correctness/security-focused, one performance/independent-correctness-focused, neither seeing the other's
output), then personally verified every high/critical claim by reading the actual source before trusting it -
one claimed bug (recurrence/location silently cleared on `RopSaveChangesMessage` update) turned out to be
backend-dependent when traced through `RepoUtils.update()` (SQL's TypeORM skips `undefined` fields; Mongo's
behavior depends on a third-party BSON default I couldn't pin down statically) - only the `title` half of that
claim was unconditionally true, so that's what got reported as CONFIRMED, and the fix (below) closes both
regardless of the backend nuance by making it correct-by-construction instead of relying on undefined-skipping.

**Fixed, in order of severity:**
- **Critical DoS**: `PropertyValue.ts`'s `readCountedArray()` looped a client-controlled `uint32` element count
  with no cap, and `BufferCursor.ts`'s `readNullTerminatedUtf16LE`/`readNullTerminatedString8` never threw at
  end-of-buffer (returned `""` and kept going) despite the class's own doc comment claiming every read "bounds-
  checks and throws" - together, one `RopSetProperties` call (reachable before the handler even validates its
  target handle) with a `PtypMultipleString` count of `0xFFFFFFFF` could hang or OOM the whole process. Both
  string readers now throw when no terminator is found before the buffer ends; `readCountedArray` independently
  rejects any count exceeding the buffer's actual remaining bytes, since every element needs ≥1 byte.
- **`readBytes()` silently clamping instead of throwing** on a negative or oversized length - `Buffer.subarray`
  doesn't throw, so a malformed `RopSize < 2` rewound `decodeRopBuffer`'s cursor backwards and mis-parsed the
  handle table instead of failing loudly. Now bounds-checked explicitly.
- **Four separate uint16-response-field overflows** that turned ordinary (not just malicious) client usage into
  a 500: `RopBuffer.ts`'s `encodeRopBuffer` writes `RopSize` as a uint16 with no cap on `ropsList.length`
  (`RopQueryRows` against a few hundred messages already exceeds it), `RopReadStreamHandler`'s `DataSize` and
  `RopFastTransferSourceGetBufferHandler`'s `TransferBufferSize` both used a 32-bit requested/remaining size
  unclamped before a 16-bit write. Fixed `RopReadStream`/`FastTransferSourceGetBuffer` by clamping to `0xFFFF`
  and letting the client page for the rest (spec-legal - real Exchange does exactly this). **Not** fixed at the
  `encodeRopBuffer` level itself (would need a shared byte-budget threaded through every ROP in an `Execute`
  batch, a materially bigger change) - `RopQueryRows` was the one call site actually exercising this, and it's
  now addressed by the same pattern as the other two if it ever needs it; flagged as a narrower remaining gap
  in that handler's own doc comment rather than silently left unaddressed.
- **`RopSetColumns` silently succeeding for a missing/wrong-type handle** (`if (handle) {...}` with no `else` -
  every sibling ROP validates and returns `MAPI_E_INVALID_OBJECT`). Now matches the sibling pattern.
- **`RopLogon` wholesale-replacing `session.folderIds`** on every call, discarding child-folder FIDs a client
  had already learned via `RopGetHierarchyTable`/`RopQueryRows` on a re-logon (spec-legal within one session) -
  and risking a re-issued FID silently pointing at a *different* folder. Fixed by having `RopLogon` assign its
  13 special-folder FIDs through the exact same `FolderTarget.assignOrGetFid` reuse-or-mint logic child folders
  already use, instead of a bespoke always-reset-to-1 counter.
- **`RopSaveChangesMessage` blanking a calendar event's title** (unconditionally, any backend) whenever an
  update touched some other property without resending `PidTagSubject` - `decodeCalendarFieldsFromDraft`
  defaulted an absent `PidTagSubject` to `""` instead of leaving it `undefined` the way every other optional
  field already did, and the update path had no existing-value fallback for it (unlike `startDate`/`endDate`/
  `busyStatus`/`timezone`, which already did). Fixed `title` to stay `undefined` when absent and added the same
  `?? existing.xxx` fallback to `location`/`recurrenceRule`/`reminderMinutesBeforeStart` too, closing the
  backend-dependent uncertainty on those by making the fallback explicit rather than relying on `update()`'s
  undefined-handling.
- **`RopDeleteFolder` not knowing about Contacts/Tasks folder content** - emptiness was checked only against
  `messageRepo`/`calendarEventRepo`, so a non-empty Contacts/Tasks folder could be deleted without `DEL_MESSAGES`
  and orphan its `Contact`/`Task` rows (a real regression from the 2026-09-08 Contacts/Tasks feature work, not
  the documented "browse-only" limitation). Both repos now participate in the emptiness check and delete
  cascade, including the recursive subfolder path - guarded the same optional-repo way `RopGetContentsTableHandler`
  already does.
- **Meeting-response correlation was unreachable by a real Outlook client, not just a documented limitation**:
  `GlobalObjectId.ts`'s own doc comment claimed this server always embeds the `GlobalObjectId` it generates into
  the invites it sends - but `RopSubmitMessageHandler.submitAppointment()`'s ICS only ever emits a bare `UID:`
  line (invites go out as plain SMTP/iCalendar, not native MAPI Store objects), and `encodeGlobalObjectId` had
  zero call sites outside its own test file. A real client therefore always *synthesizes* its own
  `PidLidGlobalObjectId` from that bare UID using the standard `"vCal-Uid"`-wrapped `VCALID` form
  (`[MS-ASEMAIL]` §2.2.2.37 / the equivalent `[MS-OXCICAL]` algorithm - verified against Microsoft's own docs
  before implementing, the same bar applied to every other spec claim in this codebase), which the codec's
  bare-UTF8-bytes decode didn't recognize. Rewrote `encodeGlobalObjectId`/`decodeGlobalObjectId` to build/parse
  the real `VCALID` wire form (falling back to raw bytes for the `OutlookID` case a native Exchange server
  produces, which this server's own SMTP-based flow never does but decodes for completeness) - this is a real
  functional fix, not just documentation, since accept/decline responses from a genuine Outlook client would
  never have correlated back to the right `CalendarEvent` before this.
- **NSPI `GetMatches` `TypeError`** on a `Contact` whose `emails` field is missing entirely (optional-chained
  the array element, not the array itself) - low real-world likelihood (`@rapidmx/restapi`'s `Contact` always
  initializes `emails: []`) but cheap to close and inconsistent with `ContactTarget.ts`'s own fully-guarded
  version of the identical lookup.
- **O(n²) session registries**: `FolderTarget.assignOrGetFid`/`MessageTarget.assignOrGetMid` did a linear scan
  of a session-lifetime-cumulative map on every single table row, and `NamedPropertyRegistry
  .assignOrGetNamedPropertyId` did the same plus a `Math.max(...spread)` id-allocation that risked
  `RangeError: Maximum call stack size exceeded` on a session that registered enough named properties (past
  V8's argument-spread limit) - the code's own prior comments claimed these scans were "never a real cost,"
  which wasn't true once you account for the registry being cumulative across a whole session's worth of table
  paging, not reset per page. All three now use an O(1) reverse-index/counter pair (`session.folderTargetIds`/
  `nextFolderId`, `messageTargetIds`/`nextMessageId`, `namedPropertyIds`/`nextNamedPropertyId` - new
  `MapiSessionContext` fields, small and purely additive to session-cache size, unlike the alternative of
  caching whole resolved entities would have been). `assignOrGetNamedPropertyId` also now returns `0x0000`
  ("unmappable," a real spec-defined value already used for `Kind = 0xFF`) once the `0x8000`-`0xFFFF` ID space
  is exhausted, instead of assigning an out-of-range id that later crashed a 16-bit write.
- **Redundant identical per-row queries**: `RopQueryRowsHandler`/`RopGetPropertiesSpecificHandler` now share one
  request-scoped `ResolutionCache` (`PropertyResolvers.ts`) across every row/column resolved in one call, so a
  hierarchy table's `hasChildren` (which re-fetched the mailbox's *entire* folder list per row via
  `resolveFolderChildren`) and a calendar table's organizer/attendee mailbox lookup each happen at most once per
  call instead of once per row. **Deliberately not fixed**: the deeper per-row `repo.findOne(uid)` N+1 for
  distinct message/calendarEvent/contact/task rows - a real batched fix needs a backend-agnostic "fetch these
  uids" query this codebase's `RepoUtils<T>` abstraction doesn't expose (Mongo's `$in` and TypeORM's `In()`
  aren't interchangeable at a call site that doesn't know which backend it's running against), and caching
  whole resolved entities in session state instead would make the already-flagged session-serialization-size
  problem (large `session.handles`/`messageIds` re-`JSON.stringify`'d on every `Execute`) significantly worse -
  documented as a known limitation in `RopQueryRowsHandler`'s own doc comment rather than attempted blind.
  Also parallelized `RopLogonHandler`'s four sequential per-special-folder repo lookups via `Promise.all`.
- Full suite: 472/472 passing (up from 447, +25 net new tests targeting each fix specifically - not just
  "still passes" coverage). 100% statement/function/line, 99.14% branch. `yarn build`/`yarn tsc --noEmit`/
  `yarn lint` clean. Not committed - left staged/unstaged per the standing commit-discipline rule.
