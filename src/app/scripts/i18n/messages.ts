import * as log from "../logger";
import { messages as messages_en } from "./messages_en";
import { messages as messages_ja } from "./messages_ja";
import { IMessageTranslation } from "./IMessageTranslation";


const availableLocales: { [key: string]: IMessageTranslation | undefined } = Object.create(null);
availableLocales.en = messages_en;
availableLocales.ja = messages_ja;

const fallbackLocale = messages_en;

export const get = (locale?: string | null): IMessageTranslation => {
    if (locale) {
        let m = availableLocales[locale];
        if (m) return m;

        log.warn(`No messages for [${locale}]`);

        if (locale.length > 2) {
            m = availableLocales[locale.substring(0, 2)];
            if (m) return m;
            log.warn(`No messages for fallback [${locale}]`);
        }
    } else {
        log.warn(`Invalid locale`);
    }

    return fallbackLocale;
}
