import { describe, expect, test } from "bun:test";
import { createHash } from "node:crypto";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { completeProfileHtml } from "../packages/pptxgenjs/src/render/completion";
import { buildCrossRuntimeDeck } from "./crossRuntimeDeck";
import { populateParityFixture } from "./visual/fixture";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

/**
 * sha256 of toHtml() for the shared decks, taken from the DOM-built HTML before #69 phase 4a
 * (ee97e5d). A refactor keeps them; a deliberate HTML change updates these in the
 * same PR, so the output change is visible in review; to see what changed, diff toHtml() against main.
 */
const EXPECTED_SHA256: Record<string, string> = {
    fixture: "c7151b9ebbfedb0557787d362db201a997bcc6ac9d0090150dd6f9343974e8e9",
    "cross-runtime": "1fbdb5018396ef36aa6268de665724032062110f3caada7638054e59477cf044",
    "wasm-text": "7e970ac08af8b3aa1479e11ba48f16aeebb10d6d79127943e3f488da6b946bd7",
    "wasm-emoji": "041e711c8a1912d6a0c73bd32aa2403839987c3929d4cbdfcb0a5115e27e8df1",
    "wasm-layout": "1a866c565c2aefbe0befd5a09d56edf9b5f9faf4e8fdeb57218ba0247704adbd",
    "wasm-shapes": "74ad079c7bdf417daf5b4f8e368ea35b487c607565958ef68bd38ba862d5027c",
    "wasm-tables": "9db006ff1efc30aadf23d6a778a023f8f365eff14875d520687566383e68a917",
    "wasm-images": "acedb8d02711e608c7933252c27af2c6b534fecb709e45d621ca170a3ef0d6ce",
    "wasm-effects": "c66bddedfa6557470ca3ff722723e60f28892ce2815af183d8afdc680a427207",
    "wasm-charts": "fb29e38a7cf3ad5c54e3bf3a888ed84a646ab24fa3b087799bc98365e9ed174c",
    "wasm-chartAxis": "a72341cda5ffdf46ac0dd0dea2a19ae6ce7040cb6aee0269df7ea28621f0ece6",
    "wasm-chartPlot": "668c51abc25242f70dd8f63b5dfe5d732d746d2fcd6cf4b19585f49355ee0bb6",
    "wasm-placeholders": "3010682db4a0a542a254c4866a24b695c4b53c879a1ad9ab5e645f6c2956a792",
    "wasm-chartMarker": "512861437e9bb3e1b0a4caa1b4bd1b09a1e76a8c2fde82fd29ccf085a28cfc76",
    "wasm-chartScatter": "b04a984ecd3496101a3f6c5f5fc06068d4bd8f195aa07986b8619010eae37252",
    "wasm-chartBarColors": "91491b8a63a878d08a16fce5f979aa7bef70d4505331e190ff702fd1c49d5efb",
    "wasm-chartPie": "b5b9ac58f88bf6a830d87b93d82135122a80cd2a87b9b367478c477a662fadea",
    "wasm-chartBarOrder": "b65d32f5288d81d4080ce7c27f579c4d277a738dca9424a418222d949fe534ad",
    "wasm-chartLegendOrder": "0fb670a22e4216b4ca25e958c671644b40e0ed5dc3f8ffca738ec51c6e278381",
    "wasm-chartStacked": "4c87d1876e0c883ebd1e829ffbf6c364824e0d3a0be06be8b79b05c1df50c529",
    "wasm-chartArea": "5b6fe73653f3901e4d6c88eb3db75ea33eefff42c166d914ae2894562ccc74f8",
    "wasm-chartPlotBand": "3c56ce5a82ad53e1698e01a34d6363d06f4cafcf5e7a09103d67911133cc899d",
    "wasm-chartCombo": "829a549da8cbbd4058f8784413e26e11c655e18e75fb0db8e74cbe8859ada863",
    "wasm-chartAreaOrder": "0dd7b361670052c66972165c2b144408162cf669dfeb2fc6c473c78022155e4b",
    "wasm-chartPointColors": "dfb77c0a549c56d85f2f2f986a9cc99ee6a65519f37163b7079ba0c7e4bfbcf2",
    "wasm-chartScatterAxes": "cccfe2aca9150f8f64e25931e78cdac2c83a1a7d0f6873c3cdf8e803ddb5837e",
    "wasm-chartTopRightLegend": "aaf39b16e26a813435a919b18fdc3fe34472d21b639f684f6f6f23fb62a7cab3",
    "wasm-chartValueAxis": "f10bcbcfa7b15b0dc6e084b343249e1cd7ba7201dbbea188eb85f8d365923073",
    "wasm-chartLabelPositions": "61abdeabe26269e349d5ce986c6f4cffc245d90aff4174918ec1d322509ce14d",
    "wasm-bulletType": "00f6f3cdb9a058f0278081965f24fb3e7663cbc96ce4b11ddff3160652ec1898",
    "wasm-chartFixedBound": "27d96ad337427a785c8d5cef3b5f9a3567956ffa404038c6361390e56d220766",
    "wasm-chartBar3DAxis": "6704d63b0b6491c788bc8cedfe9335d59019c315c6024c2cc6d85ae6f31754dc",
    "wasm-svg": "0ed2222573f3ffa6ec6b887907d3f26a2343fe30f2c4a44a4608df6e0b5fba31",
    "wasm-gif": "cc044a43bcdaa50876fc2da1fedc3d3ede74d271b7a8ce406191ddd96b4dc71e",
    "wasm-webp": "389927c7c8b1506b25d5dece837d9f0131d4af29ad94ee0b361719c07798f1b8",
    "wasm-slides": "7ceb86e4a5ef2d6629dc1766f6fd1bcaaa752d9b7fbd35b2d91fd4c915be61f2",
    "wasm-chartLabels": "2282dc2f2915e3fb504d6235bc549b54d99e950c2512232f00ce0f8663f2338d",
    "wasm-components": "306e319217049fcad4206767ddeae2b4777b4986127957b685f076fcd603f576",
    "wasm-options": "0b57e9c0b341e62b4bfcaa71dba88f9c554cc801d88a78347189590697f70da3",
    "wasm-tableOptions": "97ffa9ab07f395f5cc0c99085665bbdda51f361f20aa7f5a4e0d65ff810e168f",
};

describe("toHtml() bytes", () => {
    test("match the recorded hashes for the fixture, the cross-runtime deck and every WASM deck, and need no completion", async () => {
        const html: Record<string, string> = {};
        const fixture = new PuppeteerGen();

        populateParityFixture(fixture);
        html.fixture = (await fixture.toHtml()).html;
        html["cross-runtime"] = (await buildCrossRuntimeDeck(PuppeteerGen, PNG).toHtml()).html;
        // Imported here: tests/visual/wasm/deck.ts installs a global jsdom `document` when loaded, which
        // tests/html.test.ts must not see; it is removed again below.
        const hadDocument = "document" in globalThis;

        try {
            const { DECKS } = await import("./visual/wasm/decks");

            for (const [name, deck] of Object.entries(DECKS)) {
                html[`wasm-${name}`] = (await deck.populate(new PuppeteerGen()).toHtml()).html;
            }
        } finally {
            if (!hadDocument) {
                Reflect.deleteProperty(globalThis, "document");
            }
        }

        expect(Object.fromEntries(Object.entries(html).map(([name, document]) => [name, createHash("sha256").update(document).digest("hex")]))).toEqual(EXPECTED_SHA256);
        // toHtml() states every value, so the completion step returns it byte for byte.
        for (const [name, document] of Object.entries(html)) {
            expect({ name, same: completeProfileHtml(document) === document }).toEqual({ name, same: true });
        }
    }, 60_000);
});
