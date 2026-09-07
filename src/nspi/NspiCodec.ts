///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "../codec/BufferCursor.js";
import {
    PropertyType,
    readPropertyTag,
    readTaggedPropertyValue,
    writePropertyTag,
    writePropertyValue,
    type PropertyTag,
    type PropertyValueData,
} from "../codec/PropertyValue.js";

/**
 * Shared wire-format helpers for the NSPI-over-HTTP address-book endpoint (`[MS-OXCMAPIHTTP]` §2.2.5,
 * `[MS-OXNSPI]`) - a genuinely different transport shape from EMSMDB's ROP buffers: each `X-RequestType` value
 * (`Bind`/`Unbind`/`GetMatches`/...) has its own flat request/response body, not a multiplexed ROP stream, so
 * there is no generic framing codec to share the way `RopBuffer.ts` does for EMSMDB - just these small
 * structures every NSPI operation body embeds.
 *
 * @author Jean-Philippe Steinmetz
 */

/** `STAT` (`[MS-OXNSPI]` §2.2.8, confirmed this session): 9 `DWORD`-sized fields (36 bytes total, `Delta`
 * signed, every other field unsigned) describing an address-book container's paging/locale state. This
 * pragmatic subset never persists a `STAT` across calls (see `NspiGetMatchesHandler.ts`'s own doc comment for
 * why) - `readStat`/`writeStat` exist purely to keep the surrounding request/response body's own field
 * offsets correct, not to drive any real positioning logic. */
export interface Stat {
    sortType: number;
    containerId: number;
    currentRec: number;
    delta: number;
    numPos: number;
    totalRecs: number;
    codePage: number;
    templateLocale: number;
    sortLocale: number;
}

export function readStat(reader: BufferReader): Stat {
    return {
        sortType: reader.readUInt32LE(),
        containerId: reader.readUInt32LE(),
        currentRec: reader.readUInt32LE(),
        delta: reader.readInt32LE(),
        numPos: reader.readUInt32LE(),
        totalRecs: reader.readUInt32LE(),
        codePage: reader.readUInt32LE(),
        templateLocale: reader.readUInt32LE(),
        sortLocale: reader.readUInt32LE(),
    };
}

export function writeStat(writer: BufferWriter, stat: Stat): void {
    writer.writeUInt32LE(stat.sortType);
    writer.writeUInt32LE(stat.containerId);
    writer.writeUInt32LE(stat.currentRec);
    writer.writeInt32LE(stat.delta);
    writer.writeUInt32LE(stat.numPos);
    writer.writeUInt32LE(stat.totalRecs);
    writer.writeUInt32LE(stat.codePage);
    writer.writeUInt32LE(stat.templateLocale);
    writer.writeUInt32LE(stat.sortLocale);
}

/** A blank `STAT` (`SortType`/`ContainerID`/... all zero) - what this pragmatic subset always echoes back
 * (in `Bind`'s optional acknowledgement and `GetMatches`' response) regardless of the caller's own request
 * `STAT`, since no real per-container/locale state is tracked. */
export const BLANK_STAT: Stat = {
    sortType: 0,
    containerId: 0,
    currentRec: 0,
    delta: 0,
    numPos: 0,
    totalRecs: 0,
    codePage: 0,
    templateLocale: 0,
    sortLocale: 0,
};

/** `LargePropertyTagArray` (`[MS-OXCMAPIHTTP]` §2.2.1.8): a 4-byte count followed by that many `PropertyTag`
 * structures (`[MS-OXCDATA]` §2.9 - the same `PropertyType`-then-`PropertyId` encoding `PropertyValue.ts`'s own
 * `readPropertyTag`/`writePropertyTag` already implement). */
export function readLargePropertyTagArray(reader: BufferReader): PropertyTag[] {
    const count = reader.readUInt32LE();
    const tags: PropertyTag[] = [];
    for (let i = 0; i < count; i++) {
        tags.push(readPropertyTag(reader));
    }
    return tags;
}

export function writeLargePropertyTagArray(writer: BufferWriter, tags: PropertyTag[]): void {
    writer.writeUInt32LE(tags.length);
    for (const tag of tags) {
        writePropertyTag(writer, tag);
    }
}

/** Writes one column's `AddressBookPropertyValue` (`[MS-OXCMAPIHTTP]` §2.2.1.1): a `PtypString`/`PtypString8`/
 * `PtypBinary` value is preceded by a `HasValue` byte (always `0xFF`/`TRUE` in this pragmatic subset, which
 * never reports a missing column value - the same "always emit a real value of the right type" principle
 * `PropertyResolvers.defaultValueForType` already applies for EMSMDB); every other (fixed-size) type has no
 * such prefix. `PtypMultiple*` columns (which need a `HasValue` byte per *element*, not just once) are never
 * produced by this pragmatic subset's own fixed GAL column set and are intentionally not handled here. */
function writeAddressBookPropertyValue(writer: BufferWriter, type: PropertyType, value: PropertyValueData): void {
    if (type === PropertyType.PtypString || type === PropertyType.PtypString8 || type === PropertyType.PtypBinary) {
        writer.writeUInt8(0xff); // HasValue - always TRUE, see function doc comment
    }
    writePropertyValue(writer, type, value);
}

/** Writes one `AddressBookPropertyRow` (`[MS-OXCMAPIHTTP]` §2.2.1.7) - always `Flags=0x00` (every column
 * value present without error), the exact `StandardPropertyRow`-equivalent choice `RopQueryRowsHandler`/
 * `RopGetPropertiesSpecificHandler` already make for EMSMDB's own row encoding. */
export function writeAddressBookPropertyRow(writer: BufferWriter, columns: PropertyTag[], values: PropertyValueData[]): void {
    writer.writeUInt8(0x00); // Flags - see function doc comment
    columns.forEach((column, index) => writeAddressBookPropertyValue(writer, column.propertyType, values[index]));
}

/** Decodes a `Restriction` (`[MS-OXCDATA]` §2.12)'s leading `RestrictType` byte and, only for a
 * `ContentRestriction`/`RES_CONTENT` (`0x03`) - the shape a real client's own GAL "search as you type" sends -
 * extracts the plain string search term from its `TaggedValue` field.
 *
 * **Throws for any other restriction type** (`AND`/`OR`/`PropertyRestriction`/...) rather than degrading
 * gracefully - unlike this pragmatic subset's usual "return a default instead of failing" stance elsewhere,
 * a restriction's own byte length is type-dependent (12 different packet formats, several recursively nested),
 * so once an unrecognized `RestrictType` is seen there is no safe way to know how many bytes to skip to reach
 * the request body's own subsequent fields (`RowCount`/`Columns`/...) - silently returning `undefined` here
 * would leave the reader mid-structure and corrupt every field decoded afterward. Throwing is the honest,
 * safe choice for an input this narrow codec was never meant to parse, not a design gap.
 *
 * `FuzzyLevelLow`/`FuzzyLevelHigh` (exact/substring/prefix matching, case sensitivity) and the restriction's
 * own target `PropertyTag` are decoded to advance the reader correctly but not honored - this pragmatic
 * subset always does a case-insensitive substring match against its own fixed Contact field set regardless of
 * which property or fuzzy level the client's restriction actually named. */
export function extractContentRestrictionSearchTerm(reader: BufferReader): string {
    const restrictType = reader.readUInt8();
    if (restrictType !== 0x03) {
        throw new Error(`NspiCodec: unsupported restriction type 0x${restrictType.toString(16)} (only ContentRestriction/RES_CONTENT is supported)`);
    }
    reader.readUInt16LE(); // FuzzyLevelLow - not honored, see function doc comment
    reader.readUInt16LE(); // FuzzyLevelHigh - not honored, see function doc comment
    readPropertyTag(reader); // PropertyTag (the restriction's own target column) - not honored
    const tagged = readTaggedPropertyValue(reader);
    if (typeof tagged.value !== "string") {
        throw new Error("NspiCodec: ContentRestriction's TaggedValue was not a string.");
    }
    return tagged.value;
}
