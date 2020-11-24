import { db } from './db/Database';
import * as log from './logger';

// this must NOT be an arrow function so that [this] points to the element
const revoke = function (this: GlobalEventHandlers): void { URL.revokeObjectURL((this as HTMLImageElement).src) };

const b64toBlob = (b64Data: string, sliceSize=512) => {
    log.info(`▼▼▼ b64toBlob START b64Data: [${b64Data}] ▼▼▼`);
    let contentType = "";
    const endpos = b64Data.indexOf(";");
    if (endpos > 0) {
        const b64header = b64Data.substring(0, endpos);
        const start = b64header.indexOf(":") + 1;
        contentType = b64header.substring(start);
        log.info(`header [${b64header} start: [${start}] ==> contentType: [${contentType}]`);
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
    log.info(`▲▲▲ b64toBlob END ▲▲▲`);
    return blob;
  }

/**
 * Backing logic for custom html element <graph-image>
 * Takes a content hash in the src attribute.
 * Tries to load the corresponding image from DB.
 */
export class GraphImage extends HTMLElement {

    get src(): string | null {
        return this.getAttribute("src");
    }

    set src(url: string | null) {
        if (url) {
            this.setAttribute("src", url);
        } else {
            this.removeAttribute("src");
        }
    }

    async connectedCallback(): Promise<void> {
        const src = this.getAttribute("src");
        if (!src) return;

        const img = await db.images.get(src)
        if (!img) return;

        const data = img.dataUrl? b64toBlob(img.dataUrl?? "") : null;
        if (!data) return;

        this.removeContents();

        const el = document.createElement("img");
        el.onload = revoke;
        el.onerror = this.imageError;
        // el.src = URL.createObjectURL(img.data);
        el.src = URL.createObjectURL(data);

        this.append(el);
    }

    private removeContents() {
        while (this.lastChild) this.lastChild.remove();
    }

    private imageError = () => {
        this.removeContents();
        this.classList.add("imageError");
    }
}

if (window.customElements && window.customElements.define) {
    // define custom <graph-image> element to use the class above
    window.customElements.define('graph-image', GraphImage);
}
