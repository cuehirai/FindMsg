import { IDbEntityBase } from "../db-accessor-class-base";

/**
 * Cached graph image
 */
export interface IImageDb extends IDbEntityBase {
    // /** SHA-256 of binary image data encoded as hex string */
    // id: string;

    /** The original URL of the image */
    srcUrl: string;

    /** Timestamp when the image was downloaded */
    fetched: number;

    /** Image data */
    data: Blob;

    /** Blobから変換したBase64文字列 */
    dataUrl: string | null;

    /** Base64文字列が409,600(400KB)を超える場合にそれ以降のデータをchunkで保存 */
    dataChunk: string[];

    /** このイメージを参照しているメッセージなどのID */
    parentId: string | null;
}