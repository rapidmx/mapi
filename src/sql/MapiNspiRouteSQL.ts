///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ContactSQL, MailboxSQL } from "@rapidmx/restapi/sql";
import { BaseMapiNspiRoute } from "../BaseMapiNspiRoute.js";
import { RouteDecorators } from "@rapidrest/service-core";
const { Route } = RouteDecorators;

/**
 * SQL-backed concrete `BaseMapiNspiRoute`, mounted at `/mapi/nspi`. Exported from this plugin's
 * `./sql` entry point, so the server host mounts it without a wrapper class of its
 * own.
 *
 * @author Jean-Philippe Steinmetz
 */
@Route("/mapi/nspi")
export class MapiNspiRouteSQL extends BaseMapiNspiRoute<MailboxSQL> {
    protected mailboxClass: any = MailboxSQL;
    protected contactClass: any = ContactSQL;
}
