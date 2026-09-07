///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import type { BufferReader, BufferWriter } from "./BufferCursor.js";

/** `TypedString` (`[MS-OXCDATA]` §2.11.7) `StringType` enumeration values. */
const STRING_TYPE_NONE = 0x00;
const STRING_TYPE_EMPTY = 0x01;
const STRING_TYPE_STRING8 = 0x02;
const STRING_TYPE_REDUCED_UNICODE = 0x03;
const STRING_TYPE_UNICODE = 0x04;

/**
 * `TypedString` (`[MS-OXCDATA]` §2.11.7): a leading `StringType` byte followed by an optional string in the
 * format that byte names - used by `RopOpenMessage`'s success response (`SubjectPrefix`/`NormalizedSubject`).
 * `writeTypedString` never produces the "reduced Unicode" (`0x03`) form - that's purely a wire-size
 * optimization real Exchange applies when every character happens to fit in one byte, not something a peer is
 * ever required to send, so this codec always writes plain UTF-16LE (`0x04`) for non-empty strings. `undefined`
 * writes as "no string" (`0x00`), matching a field the spec allows to be entirely absent (e.g. no subject
 * prefix at all, as opposed to an empty one).
 */
export function writeTypedString(writer: BufferWriter, value: string | undefined): void {
    if (value === undefined) {
        writer.writeUInt8(STRING_TYPE_NONE);
    } else if (value === "") {
        writer.writeUInt8(STRING_TYPE_EMPTY);
    } else {
        writer.writeUInt8(STRING_TYPE_UNICODE);
        writer.writeNullTerminatedUtf16LE(value);
    }
}

/** Reads a `TypedString`, supporting every `StringType` value a peer could legally send (`0x00`-`0x04`) even
 * though `writeTypedString` only ever produces `0x00`/`0x01`/`0x04` itself - this side of the codec is only
 * ever exercised in this library's own round-trip tests, since a real MAPI/HTTP client never sends a
 * `TypedString` value to the server. Returns `undefined` for `StringType = 0x00` ("no string"), matching
 * `writeTypedString`'s own `undefined` input. */
export function readTypedString(reader: BufferReader): string | undefined {
    const stringType = reader.readUInt8();
    switch (stringType) {
        case STRING_TYPE_NONE:
            return undefined;
        case STRING_TYPE_EMPTY:
            return "";
        case STRING_TYPE_STRING8:
            return reader.readNullTerminatedString8();
        case STRING_TYPE_REDUCED_UNICODE: {
            let value = "";
            for (let code = reader.readUInt8(); code !== 0; code = reader.readUInt8()) {
                value += String.fromCharCode(code);
            }
            return value;
        }
        case STRING_TYPE_UNICODE:
        default:
            return reader.readNullTerminatedUtf16LE();
    }
}
