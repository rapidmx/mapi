///////////////////////////////////////////////////////////////////////////////
// Copyright (C) 2026 Jean-Philippe Steinmetz. All rights reserved.
// SPDX-License-Identifier: MPL-2.0
///////////////////////////////////////////////////////////////////////////////

/**
 * A small positional cursor over a `Buffer`, shared by every MAPI codec in this package (`MapiGuid`,
 * `PropertyValue`, `RopBuffer`, and every individual ROP handler's own request/response encoding). Unlike
 * WBXML's self-describing tag stream, MAPI's wire structures are fixed/variable-length binary records with no
 * generic "read until end tag" mechanism, so every codec needs the same little-endian primitive reads/writes
 * with automatic offset advancement - reimplementing that per file would be pure repetition across the dozens
 * of ROP handlers this phase's plan calls for, unlike EAS's WBXML codec, which only needed one such reader.
 *
 * Underlying `Buffer` methods already bounds-check and throw `RangeError` on overflow/truncation - a malformed
 * or truncated ROP buffer legitimately failing loudly is the correct behavior, not something to swallow here.
 *
 * @author Jean-Philippe Steinmetz
 */
export class BufferReader {
    private offset: number;

    public constructor(
        private readonly buffer: Buffer,
        offset = 0,
    ) {
        this.offset = offset;
    }

    public get position(): number {
        return this.offset;
    }

    public get remaining(): number {
        return this.buffer.length - this.offset;
    }

    public hasMore(): boolean {
        return this.offset < this.buffer.length;
    }

    public readUInt8(): number {
        const value = this.buffer.readUInt8(this.offset);
        this.offset += 1;
        return value;
    }

    public readInt16LE(): number {
        const value = this.buffer.readInt16LE(this.offset);
        this.offset += 2;
        return value;
    }

    public readUInt16LE(): number {
        const value = this.buffer.readUInt16LE(this.offset);
        this.offset += 2;
        return value;
    }

    public readInt32LE(): number {
        const value = this.buffer.readInt32LE(this.offset);
        this.offset += 4;
        return value;
    }

    public readUInt32LE(): number {
        const value = this.buffer.readUInt32LE(this.offset);
        this.offset += 4;
        return value;
    }

    public readFloatLE(): number {
        const value = this.buffer.readFloatLE(this.offset);
        this.offset += 4;
        return value;
    }

    public readDoubleLE(): number {
        const value = this.buffer.readDoubleLE(this.offset);
        this.offset += 8;
        return value;
    }

    public readBigInt64LE(): bigint {
        const value = this.buffer.readBigInt64LE(this.offset);
        this.offset += 8;
        return value;
    }

    public readBigUInt64LE(): bigint {
        const value = this.buffer.readBigUInt64LE(this.offset);
        this.offset += 8;
        return value;
    }

    public readBytes(length: number): Buffer {
        const value = this.buffer.subarray(this.offset, this.offset + length);
        this.offset += length;
        return value;
    }

    /** Reads a UTF-16LE string up to (and consuming) its terminating `0x0000` code unit - `PtypString`'s wire
     * encoding. Scans on 2-byte boundaries, since a null code unit's low/high byte pair can't be mistaken for
     * one half of a non-null UTF-16 code unit at an odd offset the way a naive single-byte scan could. */
    public readNullTerminatedUtf16LE(): string {
        let end = this.offset;
        while (end + 1 < this.buffer.length && !(this.buffer[end] === 0 && this.buffer[end + 1] === 0)) {
            end += 2;
        }
        const value = this.buffer.toString("utf16le", this.offset, end);
        this.offset = end + 2;
        return value;
    }

    /** Reads a single-byte-terminated 8-bit string (`PtypString8`'s wire encoding). Multibyte string content
     * itself is treated as UTF-8, a pragmatic choice documented in `PropertyValue.ts` rather than the
     * "externally specified encoding" the spec leaves open-ended. */
    public readNullTerminatedString8(): string {
        let end = this.offset;
        while (end < this.buffer.length && this.buffer[end] !== 0) {
            end += 1;
        }
        const value = this.buffer.toString("utf-8", this.offset, end);
        this.offset = end + 1;
        return value;
    }
}

/** The write-side counterpart to `BufferReader`, accumulating chunks and concatenating once via `toBuffer()` -
 * the same "array of pieces, `Buffer.concat` once at the end" idiom `WbxmlEncoder` already uses for its own
 * byte-array accumulator, adapted for multi-byte little-endian fields instead of single bytes. */
export class BufferWriter {
    private readonly chunks: Buffer[] = [];

    public writeUInt8(value: number): this {
        const buf = Buffer.alloc(1);
        buf.writeUInt8(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeInt16LE(value: number): this {
        const buf = Buffer.alloc(2);
        buf.writeInt16LE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeUInt16LE(value: number): this {
        const buf = Buffer.alloc(2);
        buf.writeUInt16LE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeInt32LE(value: number): this {
        const buf = Buffer.alloc(4);
        buf.writeInt32LE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeUInt32LE(value: number): this {
        const buf = Buffer.alloc(4);
        buf.writeUInt32LE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeFloatLE(value: number): this {
        const buf = Buffer.alloc(4);
        buf.writeFloatLE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeDoubleLE(value: number): this {
        const buf = Buffer.alloc(8);
        buf.writeDoubleLE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeBigInt64LE(value: bigint): this {
        const buf = Buffer.alloc(8);
        buf.writeBigInt64LE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeBigUInt64LE(value: bigint): this {
        const buf = Buffer.alloc(8);
        buf.writeBigUInt64LE(value, 0);
        this.chunks.push(buf);
        return this;
    }

    public writeBytes(value: Buffer): this {
        this.chunks.push(value);
        return this;
    }

    public writeNullTerminatedUtf16LE(value: string): this {
        this.chunks.push(Buffer.from(value, "utf16le"), Buffer.from([0x00, 0x00]));
        return this;
    }

    public writeNullTerminatedString8(value: string): this {
        this.chunks.push(Buffer.from(value, "utf-8"), Buffer.from([0x00]));
        return this;
    }

    public toBuffer(): Buffer {
        return Buffer.concat(this.chunks);
    }
}
