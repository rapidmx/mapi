///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ContactMongo, MailboxMongo } from "@rapidmx/restapi/mongo";
import { BaseMapiNspiRoute } from "../BaseMapiNspiRoute.js";

/**
 * Mongo-backed concrete `BaseMapiNspiRoute`. A deployment mounts this at the well-known MAPI/HTTP NSPI path via
 * its own trivial `@Route("/mapi/nspi")` subclass, following the same pattern `MapiEmsmdbRouteMongo.ts`'s doc
 * comment describes.
 *
 * @author Jean-Philippe Steinmetz
 */
export class MapiNspiRouteMongo extends BaseMapiNspiRoute<MailboxMongo> {
    protected mailboxClass: any = MailboxMongo;
    protected contactClass: any = ContactMongo;
}
