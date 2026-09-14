///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ContactMongo, MailboxMongo } from "@rapidmx/restapi/mongo";
import { BaseMapiNspiRoute } from "../BaseMapiNspiRoute.js";
import { RouteDecorators } from "@rapidrest/service-core";
const { Route } = RouteDecorators;

/**
 * Mongo-backed concrete `BaseMapiNspiRoute`, mounted at `/mapi/nspi`. Exported from this plugin's
 * `./mongo` entry point, so the server host mounts it without a wrapper class of its
 * own.
 *
 * @author Jean-Philippe Steinmetz
 */
@Route("/mapi/nspi")
export class MapiNspiRouteMongo extends BaseMapiNspiRoute<MailboxMongo> {
    protected mailboxClass: any = MailboxMongo;
    protected contactClass: any = ContactMongo;
}
