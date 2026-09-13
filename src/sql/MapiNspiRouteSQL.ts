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

    protected likePattern(term: string): string {
        // like()'s glob syntax (ModelUtils.globToLike) translates * to a SQL % wildcard - wrap on both ends for
        // a genuine substring match, the same way MapiNspiRouteMongo.ts does for its own $regex compilation.
        return `*${term}*`;
    }
}
