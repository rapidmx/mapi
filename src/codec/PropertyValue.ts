///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////
import { BufferReader, BufferWriter } from "./BufferCursor.js";
import { decodeGuid, encodeGuid } from "./MapiGuid.js";

/**
 * The property-value system every meaningful ROP depends on (`RopGetPropertiesSpecific`/`RopSetProperties`,
 * table row data in `RopQueryRows`, ...), per `[MS-OXCDATA]` §2.11. Only the property types this phase's
 * planned ROP handlers actually need are implemented (`readPropertyValue`/`writePropertyValue` throw a clear
 * error for anything else) - add more as a real handler needs them, not speculatively.
 *
 * Numeric values and member names (`Ptyp*`) are exactly as published in `[MS-OXCDATA]`'s "Property Data
 * Types" table - the full `Ptyp`-prefixed spec names are used rather than shorter aliases (`String`,
 * `Boolean`, ...) both for fidelity to the spec and because those shorter names collide with this project's
 * lint config's `id-denylist` rule (which exists to stop primitive-wrapper-like names shadowing globals).
 */
export enum PropertyType {
    PtypNull = 0x0001,
    PtypInteger16 = 0x0002,
    PtypInteger32 = 0x0003,
    PtypFloating32 = 0x0004,
    PtypFloating64 = 0x0005,
    PtypBoolean = 0x000b,
    PtypInteger64 = 0x0014,
    PtypString8 = 0x001e,
    PtypString = 0x001f,
    PtypTime = 0x0040,
    PtypGuid = 0x0048,
    PtypBinary = 0x0102,
    PtypMultipleInteger32 = 0x1003,
    PtypMultipleString8 = 0x101e,
    PtypMultipleString = 0x101f,
    PtypMultipleBinary = 0x1102,
}

export interface PropertyTag {
    propertyId: number;
    propertyType: PropertyType;
}

export interface TypedPropertyValue {
    propertyType: PropertyType;
    value: PropertyValueData;
}

export interface TaggedPropertyValue {
    propertyId: number;
    propertyType: PropertyType;
    value: PropertyValueData;
}

export type PropertyValueData = number | bigint | boolean | string | Buffer | Date | string[] | number[] | Buffer[];

/**
 * `PropertyTag` structure (`[MS-OXCDATA]` §2.9): a 4-byte value with **`PropertyType` in the low-order 16
 * bits and `PropertyId` in the high-order 16 bits** - confirmed against the spec's own bit-range description
 * and cross-checked against the well-known real property-tag constant `PR_SUBJECT = 0x0037001F` (`PidTagSubject`
 * = property ID `0x0037`, `PtypString` = type `0x001F`), which only decodes correctly with `PropertyType` as
 * the low half. On the wire (little-endian), that means `PropertyType` is read/written first.
 */
export function readPropertyTag(reader: BufferReader): PropertyTag {
    const propertyType: PropertyType = reader.readUInt16LE();
    const propertyId = reader.readUInt16LE();
    return { propertyId, propertyType };
}

export function writePropertyTag(writer: BufferWriter, tag: PropertyTag): void {
    writer.writeUInt16LE(tag.propertyType);
    writer.writeUInt16LE(tag.propertyId);
}

/** `TypedPropertyValue` structure (`[MS-OXCDATA]` §2.11.4): `PropertyType` (2 bytes) followed by the value
 * itself, encoded per `readPropertyValue`/`writePropertyValue` below. */
export function readTypedPropertyValue(reader: BufferReader): TypedPropertyValue {
    const propertyType: PropertyType = reader.readUInt16LE();
    const value = readPropertyValue(reader, propertyType);
    return { propertyType, value };
}

export function writeTypedPropertyValue(writer: BufferWriter, typed: TypedPropertyValue): void {
    writer.writeUInt16LE(typed.propertyType);
    writePropertyValue(writer, typed.propertyType, typed.value);
}

/** `TaggedPropertyValue` structure (`[MS-OXCDATA]` §2.11.4, confirmed via `RopSetProperties`'s own request-buffer
 * page): a full `PropertyTag` (4 bytes - `PropertyId` **and** `PropertyType`, unlike `TypedPropertyValue`'s bare
 * `PropertyType`) followed by the value itself. Used where a property's identity can't be inferred from
 * context (e.g. `RopSetProperties`, which sets an arbitrary, client-chosen set of properties in one call). */
export function readTaggedPropertyValue(reader: BufferReader): TaggedPropertyValue {
    const tag = readPropertyTag(reader);
    const value = readPropertyValue(reader, tag.propertyType);
    return { propertyId: tag.propertyId, propertyType: tag.propertyType, value };
}

export function writeTaggedPropertyValue(writer: BufferWriter, tagged: TaggedPropertyValue): void {
    writePropertyTag(writer, { propertyId: tagged.propertyId, propertyType: tagged.propertyType });
    writePropertyValue(writer, tagged.propertyType, tagged.value);
}

/**
 * Reads a bare `PropertyValue` (`[MS-OXCDATA]` §2.11.2) of the given, already-known `type`. `PtypBinary`'s
 * length prefix is 16 bits and every `PtypMultiple*`'s value count is 32 bits wide, per the spec's explicit
 * "in the context of ROP buffers" sizing rule - the only context every planned caller of this function
 * operates in (the wider 32-bit `PtypBinary` count only applies to extended rules / a specific MAPI/HTTP
 * structure this library doesn't implement).
 */
export function readPropertyValue(reader: BufferReader, type: PropertyType): PropertyValueData {
    switch (type) {
        case PropertyType.PtypNull:
            return 0;
        case PropertyType.PtypInteger16:
            return reader.readInt16LE();
        case PropertyType.PtypInteger32:
            return reader.readInt32LE();
        case PropertyType.PtypFloating32:
            return reader.readFloatLE();
        case PropertyType.PtypFloating64:
            return reader.readDoubleLE();
        case PropertyType.PtypBoolean:
            return reader.readUInt8() !== 0;
        case PropertyType.PtypInteger64:
            return reader.readBigInt64LE();
        case PropertyType.PtypTime:
            return filetimeToDate(reader.readBigUInt64LE());
        case PropertyType.PtypGuid:
            return decodeGuid(reader);
        case PropertyType.PtypString:
            return reader.readNullTerminatedUtf16LE();
        case PropertyType.PtypString8:
            return reader.readNullTerminatedString8();
        case PropertyType.PtypBinary:
            return reader.readBytes(reader.readUInt16LE());
        case PropertyType.PtypMultipleInteger32:
            return readCountedArray(reader, () => reader.readInt32LE());
        case PropertyType.PtypMultipleString:
            return readCountedArray(reader, () => reader.readNullTerminatedUtf16LE());
        case PropertyType.PtypMultipleString8:
            return readCountedArray(reader, () => reader.readNullTerminatedString8());
        case PropertyType.PtypMultipleBinary:
            return readCountedArray(reader, () => reader.readBytes(reader.readUInt16LE()));
        default:
            throw new Error(`PropertyValue: unsupported PropertyType 0x${(type as number).toString(16)}`);
    }
}

export function writePropertyValue(writer: BufferWriter, type: PropertyType, value: PropertyValueData): void {
    switch (type) {
        case PropertyType.PtypNull:
            return;
        case PropertyType.PtypInteger16:
            writer.writeInt16LE(value as number);
            return;
        case PropertyType.PtypInteger32:
            writer.writeInt32LE(value as number);
            return;
        case PropertyType.PtypFloating32:
            writer.writeFloatLE(value as number);
            return;
        case PropertyType.PtypFloating64:
            writer.writeDoubleLE(value as number);
            return;
        case PropertyType.PtypBoolean:
            writer.writeUInt8(value ? 1 : 0);
            return;
        case PropertyType.PtypInteger64:
            writer.writeBigInt64LE(value as bigint);
            return;
        case PropertyType.PtypTime:
            writer.writeBigUInt64LE(dateToFiletime(value as Date));
            return;
        case PropertyType.PtypGuid:
            writer.writeBytes(encodeGuid(value as string));
            return;
        case PropertyType.PtypString:
            writer.writeNullTerminatedUtf16LE(value as string);
            return;
        case PropertyType.PtypString8:
            writer.writeNullTerminatedString8(value as string);
            return;
        case PropertyType.PtypBinary: {
            const buf = value as Buffer;
            writer.writeUInt16LE(buf.length);
            writer.writeBytes(buf);
            return;
        }
        case PropertyType.PtypMultipleInteger32:
            writeCountedArray(writer, value as number[], (v) => writer.writeInt32LE(v));
            return;
        case PropertyType.PtypMultipleString:
            writeCountedArray(writer, value as string[], (v) => writer.writeNullTerminatedUtf16LE(v));
            return;
        case PropertyType.PtypMultipleString8:
            writeCountedArray(writer, value as string[], (v) => writer.writeNullTerminatedString8(v));
            return;
        case PropertyType.PtypMultipleBinary:
            writeCountedArray(writer, value as Buffer[], (v) => {
                writer.writeUInt16LE(v.length);
                writer.writeBytes(v);
            });
            return;
        default:
            throw new Error(`PropertyValue: unsupported PropertyType 0x${(type as number).toString(16)}`);
    }
}

function readCountedArray<T>(reader: BufferReader, readOne: () => T): T[] {
    const count = reader.readUInt32LE();
    const values: T[] = [];
    for (let i = 0; i < count; i++) {
        values.push(readOne());
    }
    return values;
}

function writeCountedArray<T>(writer: BufferWriter, values: T[], writeOne: (value: T) => void): void {
    writer.writeUInt32LE(values.length);
    for (const value of values) {
        writeOne(value);
    }
}

/** 100-nanosecond intervals between the FILETIME epoch (1601-01-01T00:00:00Z) and the Unix epoch
 * (1970-01-01T00:00:00Z) - `[MS-DTYP]`'s `FILETIME` structure, which `PtypTime` uses. */
const FILETIME_EPOCH_DIFF_100NS = 116444736000000000n;

/** Converts a `PtypTime` value (100-ns intervals since 1601-01-01) into a JS `Date`. Sub-millisecond
 * precision is truncated - `Date` has no finer resolution, a real, bounded, documented gap. */
export function filetimeToDate(filetime: bigint): Date {
    const millis = (filetime - FILETIME_EPOCH_DIFF_100NS) / 10000n;
    return new Date(Number(millis));
}

/** Converts a JS `Date` into a `PtypTime` value (100-ns intervals since 1601-01-01). */
export function dateToFiletime(date: Date): bigint {
    return BigInt(date.getTime()) * 10000n + FILETIME_EPOCH_DIFF_100NS;
}
