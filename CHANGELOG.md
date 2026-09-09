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
