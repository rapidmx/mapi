# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Added read-only Outlook Categories (PidNameKeywords, PS_PUBLIC_STRINGS) support for messages, resolving each Message.labelUids entry to its Label.name and returning them as PtypMultipleString via RopQueryRows/RopGetPropertiesSpecific
- Added labelRepo/labelClass as optional RopContext fields (mirroring contactRepo/taskRepo), populated by BaseMapiEmsmdbRoute
- Added end-to-end HTTP+DB integration tests for Categories in both the mongo and sql MapiEmsmdbRoute test suites
- Confirmed (via a targeted test against FolderType.ARCHIVE) that restapi's Archive folder type needs no new mapi code - generic folder/message browsing already handles any non-special-cased folder type identically

### Changed
- Bumped @rapidmx/restapi peer/dev dependency range to ^0.8.0 and @rapidrest/service-core to 2.x/^2.0.0
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Fixed
- Fixed NSPI GAL search (GetMatches) returning zero matches against restapi ^0.8.0/@rapidrest/service-core 2.x, whose like() operator changed from raw substring matching to glob syntax (* / ?) that treats an unescaped search term as an exact-match pattern - likePattern() implementations now wrap the raw term in `*...*` instead of regex-escaping it
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

## [1.0.0-beta.0] - 2026-09-09

### Added
- Added missing project files
- Added vite dev dependency
- Added read-only Contacts/Tasks folder browsing via new ContactTarget.ts/TaskTarget.ts, mirroring the existing CalendarEventTarget.ts pattern, wired through RopGetContentsTable/RopOpenMessage/RopGetPropertiesSpecific/RopQueryRows
- Added PidLidTaskStatus/PercentComplete/TaskDueDate/TaskComplete (PSETID_Task) named-property support for Task items via new TaskNamedProperties.ts
- Added PidTagGivenName/Surname/EmailAddress/BusinessTelephoneNumber/CompanyName/Title property support for Contact items
- Added PidTagDeferredSendTime ("do not deliver before") support to RopSubmitMessage, parking a deferred draft in Outbox with Message.scheduledSendTime for ScheduledSendJob to relay later
- Added PidTagReadReceiptRequested support to RopSubmitMessage, attaching a real Disposition-Notification-To header and recording Message.requestReceipt on the persisted Sent Items copy
- Added contactRepo/taskRepo/contactClass/taskClass as optional RopContext fields, populated by BaseMapiEmsmdbRoute, so existing RopContext test fixtures don't need updating
- Added end-to-end HTTP+DB integration tests for Contacts/Tasks folder browsing in both the mongo and sql MapiEmsmdbRoute test suites

### Changed
- Initial commit
- Update README's documented-gaps section to reflect the new Contacts/Tasks/deferred-send/read-receipt coverage and the deliberately-deferred Focused Inbox exposure
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>
- Change FolderTarget.assignOrGetFid/MessageTarget.assignOrGetMid/NamedPropertyRegistry.assignOrGetNamedPropertyId from an O(n) linear scan plus a Math.max(...spread) id-allocation to an O(1) reverse-index/counter pair
- Change assignOrGetNamedPropertyId to return 0x0000 once a session's entire named-property ID space is exhausted instead of assigning an out-of-range id that later crashed a 16-bit write
- Change RopQueryRowsHandler/RopGetPropertiesSpecificHandler to share one request-scoped ResolutionCache across every row/column resolved in a call, so per-row folder-list/mailbox fetches that only ever have one real answer per call happen at most once
- Change RopLogonHandler to resolve its four real-folder-type lookups concurrently instead of as four sequential round trips
- Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>

### Fixed
- Fixed @rapidmx/restapi peer dependency range (0.1.x) rejecting the ^0.3.0 releases this package is actually built and tested against
- Fixed RopSubmitMessageHandler's compose/send tests against a stale ScanPipeline mock missing the references/inReplyTo fields restapi ^0.3.0's scanAndRelay() now unconditionally reads to derive a conversation ID
- Fixed an unbounded PtypMultipleString/PtypMultipleString8 element count in RopSetProperties/NSPI GetMatches letting any authenticated caller hang or OOM the whole process with one crafted request
- Fixed BufferReader.readNullTerminatedUtf16LE/readNullTerminatedString8 silently returning truncated content past the buffer's end instead of throwing, closing the above DoS at its root
- Fixed BufferReader.readBytes() silently clamping an out-of-range or negative length instead of throwing, which let a malformed RopSize < 2 mis-parse decodeRopBuffer's handle table
- Fixed RopQueryRows/RopReadStream/RopFastTransferSourceGetBuffer throwing on any response exceeding 64KB, turning ordinary folder/message browsing and FastTransfer sync into a 500 - both now clamp to the response's real 16-bit size limit and let the client page for the rest
- Fixed RopSetColumns silently succeeding for a missing or wrong-type table handle instead of returning MAPI_E_INVALID_OBJECT
- Fixed RopLogon wholesale-resetting session.folderIds on a second logon within the same session, discarding every FID a client had already learned for a child folder
- Fixed RopSaveChangesMessage unconditionally blanking a calendar event's title (and, on at least one backend, its location/recurrence/reminder) on any update that didn't resend every property
- Fixed RopDeleteFolder checking only messageRepo/calendarEventRepo for emptiness, letting a non-empty Contacts/Tasks folder be deleted without DEL_MESSAGES and orphaning its Contact/Task rows
- Fixed meeting-response correlation being unreachable by a real Outlook client - this server's own outgoing invites never carried a PidLidGlobalObjectId, so decodeGlobalObjectId now recognizes the real "vCal-Uid"-wrapped VCALID form a client synthesizes from the plain iCalendar UID instead of assuming raw bytes
- Fixed a GAL search (NSPI GetMatches) TypeError on a Contact whose emails field is missing entirely
- Fixed changelog

### Removed
- Removed unused files

[Unreleased]: https://github.com/RapidMX/mapi/compare/v1.0.0-beta.0...HEAD
[1.0.0-beta.0]: https://github.com/RapidMX/mapi/releases/tag/v1.0.0-beta.0
