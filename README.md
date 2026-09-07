# RapidMX: MAPI

[![CI](https://github.com/RapidMX/mapi/actions/workflows/build.yml/badge.svg?branch=main)](https://github.com/RapidMX/mapi/actions/workflows/build.yml)
[![Coverage Status](https://coveralls.io/repos/github/RapidMX/mapi/badge.svg?branch=main)](https://coveralls.io/github/RapidMX/mapi?branch=main)
[![npm version](https://img.shields.io/npm/v/@rapidmx/mapi)](https://www.npmjs.com/package/@rapidmx/mapi)

MAPI over HTTP protocol support for a [`@rapidmx/restapi`](https://github.com/RapidMX/restapi)-based mail
server — covers the pragmatic subset a real Outlook desktop client (and the "New Outlook"/Monarch client, for
on-prem/hybrid mailboxes) needs: mailbox logon, folder/message browsing, compose/send, full Calendar CRUD
including meeting invites/responses, deletion, a pragmatic (full-dump, not byte-perfect ICS) incremental
sync, and a minimal NSPI address-book endpoint (`Bind`/`Unbind`/`GetMatches`) for GAL "search as you type."

Like ActiveSync, it authenticates with the same JWT the rest of a RapidREST app's routes already use — no
MAPI-specific auth code — which means a real native Outlook client needs an OAuth 2.0 Authorization Server
role in front of it to obtain one; that piece is tracked as a follow-up in `@rapidrest/auth`, not this
package.

Documented gaps: no byte-perfect ICS (a real client still works correctly against the full-dump form, just
less efficiently); no counter-proposals, meeting-forwarding, delegate scheduling, or resource-booking
auto-accept; no DST-aware timezones (fixed-offset approximation only); no recurrence exceptions; no
`RopModifyRecipients`; no public-folder support; no delegate/shared-mailbox access; no
rules/permissions/search-folder ROPs; no client-certificate enrollment; NSPI limited to
`Bind`/`Unbind`/`GetMatches` only.

A [`@rapidmx/autodiscover`](https://github.com/RapidMX/autodiscover) mount lets real Outlook clients find
this package's endpoints from just an email address.

## Usage

Mount `MapiEmsmdbRouteMongo`/`SQL` (mailbox/store access) and `MapiNspiRouteMongo`/`SQL` (address book) from
`@rapidmx/mapi/mongo` or `/sql` at their conventional paths, each with a trivial one-line subclass:

```ts
import { MapiEmsmdbRouteMongo, MapiNspiRouteMongo } from "@rapidmx/mapi/mongo";
import { RouteDecorators } from "@rapidrest/service-core";
const { Route } = RouteDecorators;

@Route("/mapi/emsmdb")
export class MyMapiEmsmdbRoute extends MapiEmsmdbRouteMongo {}

@Route("/mapi/nspi")
export class MyMapiNspiRoute extends MapiNspiRouteMongo {}
```

Requires a [`@rapidmx/restapi`](https://github.com/RapidMX/restapi)-backed application (models, blob
storage, and scan pipeline).

## Status

Complete. This package was carved out of the former `@rapidrest/mail` monolith — see `.claude/NOTES.md` for
the split's own rationale and history.
