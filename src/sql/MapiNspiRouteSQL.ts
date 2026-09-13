///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { ContactSQL, MailboxSQL } from "@rapidmx/restapi/sql";
import { BaseMapiNspiRoute } from "../BaseMapiNspiRoute.js";

/**
 * SQL-backed concrete `BaseMapiNspiRoute`. See `MapiNspiRouteMongo.ts`'s doc comment - the same mounting
 * pattern applies here.
 *
 * @author Jean-Philippe Steinmetz
 */
export class MapiNspiRouteSQL extends BaseMapiNspiRoute<MailboxSQL> {
    protected mailboxClass: any = MailboxSQL;
    protected contactClass: any = ContactSQL;
}
