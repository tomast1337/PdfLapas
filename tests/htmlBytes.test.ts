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
    fixture: "b2de58485d40fb9f92f1fd82f42974364c39f9138dc5cf48c8bfda3aa8e75af2",
    "cross-runtime": "798182f82b52c487635c8b55eeb4f84ea1f3ea8f596f2f0d00a092b1a9fa9e3a",
    "wasm-text": "058a6e35a65d379eb40947b3ce9efc91363b985d3a61b739c243ee5bc8009e5c",
    "wasm-emoji": "5bf811cc1df11df761adc4a141badd8a8ac5af271b30af4141834160fc71eac1",
    "wasm-layout": "5fed0cde680e2902121480cb94b3427fe7cb4e5a3ce836df4a708da288e343b5",
    "wasm-shapes": "072352f989a8d088b0db27db4cf129dd7fb336a4552063c8106f71bd16c06d7a",
    "wasm-tables": "df9de65bc0f82cb4c773b616064d58861f5ed9f8d5d0c034b519de2c2a57f398",
    "wasm-images": "51323db1fbab371aedb7c05f7c3a82a8c4116eb55c82785325dd0d0c55f0f48d",
    "wasm-effects": "a14a6b32feec18ba59e7dd6826e860973a0f1bcbbb4430a64010a612e273859e",
    "wasm-charts": "206259221a266be0ee5119fc4f19453df382a10d548b915763615d77dc7b23cd",
    "wasm-chartAxis": "f216ad1061b68f1e5488e08c6247e26200cbc2e83be5c1982ae714e5fd2fed6b",
    "wasm-chartPlot": "dd486add0911572e240877f387ee0d4ed5c310aeb5aea2f1e2800bf7b73be9a5",
    "wasm-placeholders": "4be4f3bbe4ae7a73633a42c4f1e1b3fbb7d8d4fd63c48ebcfb0c4a964971cc31",
    "wasm-chartMarker": "63a16993f8738fc31a21801edb2f09ca1bbe42238bcd002ffb60cc171b36e4a6",
    "wasm-chartScatter": "2864f9e92693c1a0c804a002ef96fc1b236c67a12ce6ca09b069ccaa55bc9923",
    "wasm-chartBarColors": "af0b62ba21b448269c408dc1a6a15c6ca599b53df392e341941857e79ed11159",
    "wasm-chartPie": "92c1c56c01dad1fbb95638f3b7d381b2cd904cedf45a0d7e2b14010760a045e0",
    "wasm-chartBarOrder": "6f6e16cc58229003c1a56de48a8b5b495f5c5c459314aef2289186632b7d35ef",
    "wasm-chartLegendOrder": "8ace4b375c57c8c62e3ce0be241a548da611dfe099ad26289b02558833827254",
    "wasm-chartStacked": "4dbf5912f915a22e5641436d204e8419bd1e18cdf68ee6905cdd4599f151c8b0",
    "wasm-chartArea": "90401954d65620d3a64edb0aefc2543457c14f81f0f54f5157b037e8bbb1312a",
    "wasm-chartPlotBand": "ac09e26da96bbba897bc9bbd08a8f63edd71989fa7ad23a9f1ff6e2c86198abf",
    "wasm-chartCombo": "0af7571759972adf7ff8ef262b34ccfce2fdc9fd030b01ddba5e98a8bd7093a8",
    "wasm-chartComboAreaPieLabels": "b578aff5cc1d0049823b2a3cdb56db5d13c53039d3b41e7cf1d2355bf36d99eb",
    "wasm-chartAreaOrder": "022085ce7727a4d491c1ff6d15a979efc8b1a74953a68c4553918d1b8c4f7b9d",
    "wasm-chartPointColors": "032f4821a1cbc2c8074821154d251bde7fd83727e837d22d01ebe74e4d62a2aa",
    "wasm-chartScatterAxes": "470bc3bb887b9a39b1c1e5476453ae21f2204f760ac62438d3841a08b211ba06",
    "wasm-chartTopRightLegend": "42cc8234168227b6401917497d868a442abfb0b5aa363128021c0c72816046b4",
    "wasm-chartValueAxis": "ac33dd444b145af021b9e0c17ecf6cc42137ee02aa45164a26cab1bf1d30a5e4",
    "wasm-chartLabelPositions": "e567bc294ae53acc73cb2515b693bf47e03c8f12aeff7192ee33898fd2af40e5",
    "wasm-bulletType": "f179ca12e2662745c96cabb36686581f5aee091477903004965376b5d947b81d",
    "wasm-chartFixedBound": "4d4a46296dfa212f745d22397f1b6d1eb47dfa08bc841a3acc487595aaf1059d",
    "wasm-chartBar3DAxis": "39b680bac06e76ff7fece6085eeac91006e8df8040a2e7937f8004d03bdea97e",
    "wasm-chartOptionLabels": "5ee65e47b4aecd8d7cbc4d8e26dcbf36c3aa7ca4b68639e89454c1cdc3659b9b",
    "wasm-chartOptionAxes": "e88b78a5a99b11d8d99c4f4a20943ee40ae1ad542587e78b0ba791135432a574",
    "wasm-chartCategoryOptions": "7564b073471fbf1b01720e5edf9ed0d7c60cc4d677e1691e4268c13579f583ad",
    "wasm-chartDataTable": "9276f0bf56e1f8386dc331053b888b3f379321bc1464f375978752ebe601495e",
    "wasm-chartPieBestFit": "dadb88f9e7e19a937be63cb903de54bfcd4cb07374e0c8511eb61bce2e446a6c",
    "wasm-chartAutomaticScale": "ccaca792aca5fe8cdf8b99eac23ff6d9f57270e4ca77b29b3891610e25e13d1d",
    "wasm-chartDateAxis": "b7720655fd8d41def96490dcd3d4cf9e2b37dd962df77b45f8926090ba002757",
    "wasm-chartTypeOptions": "ad6dc22bc9b891716e9a9ba03249c7845b0ca287f8efb497acc712efd5b19286",
    "wasm-chartOptionSources": "a7a67c7b7a6f7c9954041189b0b977ecd7939e1e7592d5f7adfc41b3edc75500",
    "wasm-fallbackFaces": "12badd5c5c44a2bf825c3e031c84f17770c4bc967a9c0dd1c3657ec5eca1f793",
    "wasm-svg": "2b26eef93b93cb4db70c129033a2254abe58392fa07808035b984ad00f9b16b6",
    "wasm-gif": "3330a509715781164e66e315207541f4cfa74d529408a464299a269cb5ca14c2",
    "wasm-webp": "eef239f48420c2abef61d0f75296177d76c3b05ac10947ff84389fe330f6fd2d",
    "wasm-slides": "44dd0d9d8785331a34e157c7db436619780a3ed9aa2e246e559e0d02d6653fa5",
    "wasm-chartLabels": "c0c91524ad473cb225bc867f2542af3ca7ca1da4ecd7db354f0483eeb129d7c1",
    "wasm-components": "8a3992fb2289ec6bfe88a1f0670087e65a751ac3bdf283e079842c143bdd62af",
    "wasm-options": "586236bc73da192cafabdcd96678a63e48a29b336f6bb6b83a4fb23f3fb96c15",
    "wasm-tableOptions": "718d68643d1b5e66b593d0e7c16fda0e9b51822753a79d40e23ffe391446f853",
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
