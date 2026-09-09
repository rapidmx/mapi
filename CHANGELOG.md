# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Added read-only Contacts/Tasks folder browsing (`RopGetContentsTable`/`RopOpenMessage`/`RopGetPropertiesSpecific`/`RopQueryRows` now resolve `"contact:"`/`"task:"` rows via new `ContactTarget.ts`/`TaskTarget.ts`, mirroring the existing Calendar target pattern), backed by `@rapidmx/restapi`'s `Contact`/`Task` models
- Added `PidLidTaskStatus`/`PercentComplete`/`TaskDueDate`/`TaskComplete` (`PSETID_Task`) named-property support for Task items, approximated from `Task.completed` (no in-progress/waiting/deferred tracking in this data model)
- Added `PidTagGivenName`/`Surname`/`EmailAddress`/`BusinessTelephoneNumber`/`CompanyName`/`Title` property support for Contact items
- Added `PidTagDeferredSendTime` ("do not deliver before") support to `RopSubmitMessage`, parking a deferred draft in Outbox with `Message.scheduledSendTime` set for `ScheduledSendJob` to relay later, mirroring `BaseMessageRoute.send()`'s own identical branch
- Added `PidTagReadReceiptRequested` support to `RopSubmitMessage`, attaching a real `Disposition-Notification-To` header and recording `Message.requestReceipt` on the persisted Sent Items copy

### Fixed
- Fixed the `@rapidmx/restapi` peer dependency range (`0.1.x`) rejecting the `^0.3.0` releases this package is actually built and tested against
- Fixed `RopSubmitMessageHandler`'s compose/send tests against a stale `ScanPipeline` mock missing the `references`/`inReplyTo` fields `@rapidmx/restapi` ^0.3.0's `scanAndRelay()` now unconditionally reads to derive a conversation ID
- Fixed an unbounded `PtypMultipleString`/`PtypMultipleString8` element count in `RopSetProperties`/NSPI `GetMatches` letting any authenticated caller hang or OOM the whole process with one crafted request - `BufferReader.readNullTerminatedUtf16LE`/`readNullTerminatedString8` now throw instead of silently returning `""` past the buffer's end, and `readCountedArray` independently rejects a count larger than the buffer could possibly hold
- Fixed `BufferReader.readBytes()` silently clamping an out-of-range or negative length instead of throwing, which let a malformed `RopSize < 2` mis-parse `decodeRopBuffer`'s handle table instead of failing loudly
- Fixed `RopQueryRows`/`RopReadStream`/`RopFastTransferSourceGetBuffer` throwing (500ing the whole request) whenever a single response exceeded 64KB - ordinary folder/message browsing, long message bodies, and FastTransfer sync all clamp to the response's real 16-bit size limit and let the client page for the rest, instead of building an oversized response and failing to encode it
- Fixed `RopSetColumns` silently succeeding for a missing or wrong-type table handle instead of returning `MAPI_E_INVALID_OBJECT`, which surfaced the real error one ROP later and behaved differently from every sibling handler
- Fixed `RopLogon` wholesale-resetting `session.folderIds` on a second logon within the same session, discarding every FID a client had already learned for a child folder (and risking a re-issued FID silently pointing at a different folder) - re-logon now reuses `FolderTarget.assignOrGetFid`'s existing-target-aware assignment instead of a fresh 1-13 renumbering
- Fixed `RopSaveChangesMessage` unconditionally blanking a calendar event's title (and, on at least one backend, its location/recurrence/reminder) whenever an update touched some other property (e.g. just `PidLidBusyStatus`) without resending `PidTagSubject` - every field an update didn't touch now explicitly falls back to the existing row's own value, matching the handler's own doc comment
- Fixed `RopDeleteFolder` checking only `messageRepo`/`calendarEventRepo` for emptiness, letting a non-empty Contacts/Tasks folder be deleted without `DEL_MESSAGES` and orphaning its `Contact`/`Task` rows - both repos are now included in the emptiness check and the delete cascade, including the recursive subfolder path
- Fixed `MeetingMessageClassHandler`'s meeting-response correlation being fundamentally unreachable by a real Outlook client: this server's own outgoing meeting invites never actually carried a `PidLidGlobalObjectId` (only a plain iCalendar `UID:` line), so a real client always synthesizes its own `GlobalObjectId` from that bare UID via the standard `"vCal-Uid"`-wrapped form (`[MS-ASEMAIL]`/`[MS-OXCICAL]`), which `decodeGlobalObjectId` didn't recognize - it now unwraps that form (falling back to the original raw-bytes decoding for a native-Exchange-style `OutlookID`, which this server's own SMTP-based invite flow never produces but decodes for completeness)
- Fixed a GAL search (`NSPI GetMatches`) `TypeError` on a `Contact` whose `emails` field is missing entirely (not just empty), which would 500 the whole search instead of degrading that one contact's email column to an empty string

### Changed
- Changed `RopQueryRowsHandler`/`RopGetPropertiesSpecificHandler` to share one request-scoped `ResolutionCache` across every row/column they resolve in a single call, so a hierarchy table's `hasChildren` column and a calendar table's organizer/attendee resolution each fetch their (call-wide-constant) data at most once instead of once per row
- Changed `FolderTarget.assignOrGetFid`/`MessageTarget.assignOrGetMid`/`NamedPropertyRegistry.assignOrGetNamedPropertyId` from an O(n) linear scan (plus, for the first two, a `Math.max(...spread)` id-allocation that risked a stack overflow on a large enough mailbox/session) to an O(1) reverse-index/counter lookup - these registries only ever grow for a session's lifetime, so the old scan-per-lookup pattern was quadratic over a session that pages through a large mailbox
- Changed `RopLogonHandler` to resolve its four real-folder-type lookups (Inbox/Outbox/Sent Items/Deleted Items) concurrently instead of as four sequential round trips
- Changed `assignOrGetNamedPropertyId` to return `0x0000` ("unmappable", the same value `[MS-OXCPRPT]` already defines for a `Kind = 0xFF` entry) once a session's entire `0x8000`-`0xFFFF` named-property ID space is exhausted, instead of assigning an out-of-range ID that later threw when written back as a 16-bit field
