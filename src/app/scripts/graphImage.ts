import { ImageTable } from './db/Image/ImageEntity';

// this must NOT be an arrow function so that [this] points to the element
const revoke = function (this: GlobalEventHandlers): void { URL.revokeObjectURL((this as HTMLImageElement).src) };

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

        // const img = await db.images.get(src)
        const img = await ImageTable.get(src)
        if (!img) return;

        // const data = img.dataUrl? b64toBlob(img.dataUrl?? "") : null;
        // if (!data) return;

        this.removeContents();

        const el = document.createElement("img");
        el.onload = revoke;
        el.onerror = this.imageError;
        el.src = URL.createObjectURL(img.data);
        // el.src = URL.createObjectURL(data);

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
