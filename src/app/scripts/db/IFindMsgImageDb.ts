/**
 * Cached graph image
 */
export interface IFindMsgImageDb {
    /** SHA-256 of binary image data encoded as hex string */
    id: string;

    /** The original URL of the image */
    srcUrl: string;

    /** Timestamp when the image was downloaded */
    fetched: number;

    /** Image data */
    data: Blob;

    /** Blobから変換したBase64文字列 */
    dataUrl: string | null;
}