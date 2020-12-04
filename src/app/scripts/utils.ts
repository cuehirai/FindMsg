import { info, warn, error } from "./logger";
import * as Bowser from "bowser";


/**
 * Ensures that the parameter is not undefined or null
 * @param value
 */
export function assert<T>(value: T | null | undefined, name = "value"): T | never {
    if (value === undefined || value === null) {
        throw new Error(`${name} must not be null or undefined`);
    }

    return value;
}


/**
 * Ensures that the parameter is of the specified type
 * @param value
 * @param type result of typeof operator or a check function that returns true when value is of the correct type
 */
// eslint-ignore-next-line @typescript-eslint/explicit-module-boundary-types
export function assertT<T>(value: unknown, type: string | ((maybeT: unknown) => boolean), name = "value"): T | never {
    if (typeof type === "string" && typeof value !== type) {
        throw new Error(`${name} is not of type ${type}`);
    }

    if (typeof type === "function" && !type(value)) {
        throw new Error(`${name} is not of the expected type`);
    }

    return value as T;
}


/**
 * Ensures that the parameter is not undefined or null in a way typescript can detect
 * @param value
 * @param name
 */
export function assert1<T>(value: T | null | undefined, name = "value"): asserts value is T {
    if (value === undefined || value === null) {
        throw new Error(`${name} must not be null or undefined`);
    }
}


export function filterNull<T>(value: T | null): value is Exclude<T, null> {
    return value !== null;
}

export declare type cancelFn = () => void;
export declare type throwFn = () => void | never;
export declare type checkFn = () => boolean;
export declare type progressFn = (status: string) => void;

/**
 * A function that does nothing
 */
export function nop(): void { /* empty */ }


export class OperationCancelled extends Error {
    constructor() {
        super("Cancelled by user request");
    }
}


/**
 * Enables cancellation similar to .NET CancellationToken
 */
export function cancellation(): [cancelFn, throwFn] {
    let cancelled = false;

    function cancel() {
        cancelled = true;
    }

    function throwIfCancellationRequested() {
        if (cancelled) throw new OperationCancelled();
    }

    return [cancel, throwIfCancellationRequested];
}


/**
 * Enables cancellation similar to .NET CancellationToken
 */
export function cancellationNoThrow(): [cancelFn, checkFn] {
    let cancelled = false;

    function cancel() {
        cancelled = true;
    }

    function isCancellationRequested() {
        return cancelled
    }

    return [cancel, isCancellationRequested];
}


/**
 * Convert object to string using JSON.stringify()
 * @param o
 */
export function safeStringify(o: unknown): string {
    if (typeof o === "string") {
        return o;
    }

    const seen = new WeakSet();
    const replacer = (_: unknown, value: unknown) => {
        if (typeof value === 'object' && value !== null) {
            if (seen.has(value)) {
                return "[already seen]";
            }
            seen.add(value);
        }
        return value;
    };

    try {
        return JSON.stringify(o, replacer, 4);
    } catch (error) {
        return "[safeStringify: error]";
    }
}


/**
 * Return a promise that resolves after the specified amount of time
 * Semantics of cancellation:
 * If cancel() throws, delay will reject with OperationCancelled
 * If cancel() returns true, delay will resolve
 * @param ms time in miliseconds
 */
export const delay = (ms: number, cancel?: checkFn | throwFn): Promise<void> => new Promise((resolve, reject) => {
    let timeout = -1;
    let interval = -1;
    let done = false;

    const cleanup = () => {
        if (done) {
            if (timeout >= 0) clearTimeout(timeout);
            if (interval >= 0) clearInterval(interval);
            timeout = interval = -1;
        }
    };

    if (cancel) {
        const checkCancel: TimerHandler = (): boolean => {
            try {
                if (cancel()) {
                    done = true;
                    resolve();
                    return true;
                }
            } catch (error) {
                done = true;
                reject(new OperationCancelled());
                return true;
            } finally {
                cleanup();
            }
            return false;
        };

        if (checkCancel()) return;
        interval = setInterval(checkCancel, 50);

    }

    const handler: TimerHandler = () => {
        done = true;
        resolve();
        cleanup();
    };

    timeout = setTimeout(handler, ms);
});


/**
 * Fix the message link
 * Message links returned from ms graph have ':' and '@' in the channelId encoded as %3A and %40.
 * Teams will not recognize the channel like this, so decode it.
 * @param url
 */
export function fixMessageLink(url: string): string {
    if (typeof url !== "string" || url.trim() === "") return "";

    try {
        const a = new URL(url);
        a.pathname = decodeURIComponent(a.pathname);
        return a.toString();
    } catch (err) {
        warn("invalid url", err);
        return "";
    }
}


/**
 * Checks if argument is an object and narrows type
 * @param x
 */
export function isUnknownObject(x: unknown): x is { [key in PropertyKey]: unknown } {
    return x !== null && typeof x === 'object';
}


/*
Problem statetemt:
If we don't have the persistent storage permission, the browser is free to delete out indexDBs any time it likes.
So we need this permission. But there are problems.

- If we are inside a cross-origin iframe, permission requests tend to be silently refused.
- In chrome, there is no UI for asking for permission. A certain amount of user interaction is required. We still have to call persist()!?
  This also seems to apply to other chromium based browsers.
- In firefox, we need to ask the user
- Behaviour of Safari is unknown

What to do about it:
- Show an unobtrusive indicator when permission is not granted
- I firefox, display a widget to call the user to action
*/
export const storage = (() => {
    const p = Bowser.getParser(navigator.userAgent);
    const name = p.getBrowserName(true);

    const isRunningInFirefox = name === "firefox";
    const isStorageManagerAvailable = navigator.storage && typeof navigator.storage.persist === "function" && typeof navigator.storage.persisted === "function";

    const inIFrame = (function inIframe() {
        try {
            return window.self !== window.top;
        } catch (e) {
            return true;
        }
    })();

    let permissionGranted = true;

    (async function check() {
        if (!isStorageManagerAvailable) return;
        if (await navigator.storage.persisted()) return;

        permissionGranted = false;

        while (!(permissionGranted = await navigator.storage.persisted())) {
            if (!isRunningInFirefox) {
                permissionGranted = await navigator.storage.persist();
            }
            await delay(20000);
        }
    })().catch(error);

    return Object.freeze({
        granted: () => permissionGranted,
        askForPermission: isStorageManagerAvailable && isRunningInFirefox,
        needNewWindow: inIFrame && isRunningInFirefox,
    });
})();


/** match the URL of an image attached to a channel message */
const channelMsgContentRex = /^https:\/\/graph\.microsoft\.com\/.*\/teams\/[a-zA-Z0-9-]{36}\/channels\/[^/]+\/messages\/[0-9]+\/(replies\/[0-9]+\/)?hostedContents\//i.compile();
/** match the URL of an image attached to a chat message */
const chatMsgContentRex = /^https:\/\/graph\.microsoft\.com\/.*\/chats\/[^/]+\/messages\/[0-9]+\/hostedContents\//i.compile();

/**
 * Return true, if the URL points to a message hosedContent on ms graph
 * @param url
 */
export function isGraphHostedContentUrl(url: string): boolean {
    try {
        new URL(url);
    } catch (error) {
        return false;
    }

    return channelMsgContentRex.test(url) || chatMsgContentRex.test(url);
}


/**
 * Converts a Blob to an ArrayBuffer
 * (because Safari does not support blob.arrayBuffer())
 * @param blob
 */
function blob2arrayBuffer(blob: Blob): Promise<ArrayBuffer> {
    if (blob.arrayBuffer) {
        return blob.arrayBuffer();
    } else {
        const fr = new FileReader();
        fr.readAsArrayBuffer(blob);

        return new Promise<ArrayBuffer>(function (resolve, reject) {
            fr.onload = fr.onerror = function (evt) {
                fr.onload = fr.onerror = null;

                if (evt.type === 'load' && fr.result instanceof ArrayBuffer) {
                    resolve(fr.result);
                } else {
                    reject(new Error('Failed to read the blob'))
                }
            }
        });
    }
}


/**
 * Generate a SHA-256 hash the content of the blob
 * @param content the blob to hash
 * @returns SHA-256 hash as hex string
 */
export async function hashBlob(content: Blob): Promise<string> {
    const buffer = await blob2arrayBuffer(content);
    const res = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', buffer)));
    return res.map(b => b.toString(16).padStart(2, '0')).join('');
}

/**
 * BlobをDataUrlに変換
 * @param blob 
 */
export async function blob2dataUrl(blob: Blob): Promise<string> {
    const fr = new FileReader();
    fr.readAsDataURL(blob);
 
    return new Promise<string>(function (resolve, reject) {
        fr.onload = fr.onerror = function (evt) {
            fr.onload = fr.onerror = null;
 
            if (evt.type === 'load' && typeof fr.result == 'string') {
                resolve(fr.result);
            } else {
                reject(new Error('Failed to read the blob'))
            }
        }
    });
 }
 
/**
 * base64文字列からBlobを生成します。
 * @param b64Data FileReader.readAsDataURLでBlobから変換したbase64文字列
 * @param sliceSize 変換する際に入力をスライスするサイズ※パフォーマンスに影響するらしい(512くらいがベストだそう)
 */
export function b64toBlob(b64Data: string, sliceSize=512):Blob {
    info(`▼▼▼ b64toBlob START b64Data: [${b64Data}] ▼▼▼`);
    let contentType = "";
    const endpos = b64Data.indexOf(";");
    if (endpos > 0) {
        const b64header = b64Data.substring(0, endpos);
        const start = b64header.indexOf(":") + 1;
        contentType = b64header.substring(start);
        info(`header [${b64header} start: [${start}] ==> contentType: [${contentType}]`);
    }
    const byteCharacters = atob(b64Data.replace(/^.*,/, ""));
    const byteArrays: BlobPart[] = [];
  
    for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
      const slice = byteCharacters.slice(offset, offset + sliceSize);
  
      const byteNumbers = new Array(slice.length);
      for (let i = 0; i < slice.length; i++) {
        byteNumbers[i] = slice.charCodeAt(i);
      }
  
      const byteArray = new Uint8Array(byteNumbers);
      byteArrays.push(byteArray);
    }
  
    const blob = new Blob(byteArrays, {type: contentType});
    info(`▲▲▲ b64toBlob END ▲▲▲`);
    return blob;
}

