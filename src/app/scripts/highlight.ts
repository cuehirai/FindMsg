import Mark from 'mark.js/src/lib/mark';

const markEl = "MARK";

export const noHighlight = ["", ""] as Readonly<[string, string]>;

/** Collapse only text nodes longer than this number of characters */
const collapseThreshDefault = 120;
/** Leave this many characters before and after a highlight as context */
const collapseContextDefault = 40;
/** Insert this string in place of collapsed characters */
const collapseEllipsis = " [...] ";
/** Ignore these html elements when collapsing */
const collapseBlackList = new Set(["AUDIO", "BR", "BUTTON", "CANVAS", "EMBED", "IFRAME", "IMG", "INPUT", "OBJECT", "SCRIPT", "SELECT", "STYLE", "SVG", "TEMPLATE", "VIDEO"]);

/*
Note:
mark.js works asynchronously, so technically, the completely impure way of returning if there are any hightlights would not work.
However, I think that it is actually working synchronously as long as no iframes are involved.
This way, I believe we might save hundreds if object creations and awaits.
*/
let matches = 0;

/** Options for mark.js. For full list of possible options see https://markjs.io/ */
const markOptionsSearch = {
    element: markEl,
    className: "search",
    separateWordSearch: false,
    diacritics: false,
    done: (totalMatches: number) => matches += totalMatches,
};

const markOptionsFilter = {
    element: markEl,
    className: "filter",
    separateWordSearch: false,
    diacritics: false,
    acrossElements: true,
    done: (totalMatches: number) => matches += totalMatches,
};


/**
 * Deeply iterate over node and collapse all text nodes whose text ist longer then collapseTresh.
 * Parts that are highlighted by mark.js are left untouched.
 * @param node the target noce
 * @param thresh collapse only text longer than this many characters
 * @param context leave this many characters standing before and after ellipsis
 * @returns number of characters that were collapsed
 */
export function collapse(node: HTMLElement, thresh = collapseThreshDefault, context = collapseContextDefault): number {
    if (context * 2 >= thresh) return 0;
    const stat = { elided: 0 };
    collapseImpl(node, thresh, context, stat);
    return stat.elided;
}


function collapseImpl(node: HTMLElement, collapseThresh: number, collapseContext: number, o: { elided: number }): void {
    /*
      Implementation Note:
      Iterate over the nodes recursively.
      If node is <mark> with attribute data-markjs="true" or otherwise blacklisted
      -> skip
      -> else recurse
      If textnode, check length
      -> if len <= thresh, skip
      -> if len >  thresh, replace with   <firstFewWords><...><lastFewWords>
      Keep stats how many chars were omitted
      Return total number of chars collapsed (not counting inserted omission markers)
    */

    for (let target = node.firstChild; target !== null; target = target.nextSibling) {
        switch (target.nodeType) {
            case Node.ELEMENT_NODE: {
                const el = target as HTMLElement;
                if (collapseBlackList.has(el.tagName)) continue;
                if (el.tagName === markEl && el.dataset.markjs === "true") continue;
                collapseImpl(el, collapseThresh, collapseContext, o);
                break;
            }
            case Node.TEXT_NODE: {
                const tn = target as Text;
                const len = tn.textContent?.length ?? 0;
                if (len > collapseThresh) {
                    const end = tn.splitText(len - collapseContext);
                    const ellipsis = tn.splitText(collapseContext);
                    ellipsis.textContent = collapseEllipsis;
                    target = end;
                    o.elided += len - 2 * collapseContext;
                }
                break;
            }
            default: continue;
        }
    }
}


/** Check if the array contents are strictly equal */
export function highlightEqual(a: Readonly<[string, string]>, b: Readonly<[string, string]>): boolean {
    return (a === b) || !((a[0] !== b[0]) || (a[1] !== b[1]));
}


/** Hightlight content of this node */
export function highlightNode(node: Node, [search, filter]: [string, string]): boolean {
    if (search || filter) {
        const mSub = new Mark(node);
        matches = 0;
        if (search) mSub.mark(search, markOptionsSearch);
        if (filter) mSub.mark(filter, markOptionsFilter);
        return matches > 0;
    } else {
        return false;
    }
}


/** Remove all children */
export function empty(node: Node): void {
    while (node.lastChild) node.lastChild.remove();
}
