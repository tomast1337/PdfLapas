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
    fixture: "d1c8ebfacc26def7eef46e3d3283e5c658718b661f26b0266e82f7f3686dbb20",
    "cross-runtime": "f91873120d39362618c9a6e5fec29d9353be09d25dd9c718bb7af8dfc7e7a9fd",
    "wasm-text": "9700fb8fe52842ff7f6f09e6cd4c3c5b3a9af6d01659ab00ab124f2d3c479641",
    "wasm-emoji": "0170f6d0e30f8fb4961eda77afac9a9ab91bd91534aaa2c394c68edd9d332fc0",
    "wasm-layout": "0b7400fe7c531e7e1a092763ab73f5ed77d2a19390322c20206bdde835708ec9",
    "wasm-shapes": "aff5c6700f8b6f4b00127e40e8a8b2412ca821e06c8f233ea001a2596720be3f",
    "wasm-tables": "0cbe750637e62232cca8447e143d6166bcd1861e6f3f19aa76c44cd4f17b4a46",
    "wasm-images": "119660443b85adcf0dbeeaa887c4112b27efb546e49678a4c50cb5863e25dd16",
    "wasm-effects": "5c44ac688170f0772061d8f4a7c5ba3359c73947123207deb16e61298bd30577",
    "wasm-charts": "88affd22d46387693e2eb6f10836d01a763735d26b7fb4a5755e53486fe4f3fd",
    "wasm-chartAxis": "a80e7e6e7b3581711fc9b80a5cd20574d676b282914b1e54939fe39a39a01c34",
    "wasm-chartPlot": "b3d6edd9e6d43821a15ae72d6c20103842a164c2dbdaad0bb6f18f57eacdeed5",
    "wasm-placeholders": "53821eb8897ea93c67c9739a6c027813277946713717d6a4c7c138d19daf0810",
    "wasm-chartMarker": "7c4947e80e6e89f4e93ea99cde8a0e2ec95725499751c9c01bab44ea269c1df4",
    "wasm-chartScatter": "5a5678bacb4457e806824c1256daa13ceb1ad73da04cb110c3e2330fff3677c8",
    "wasm-chartBarColors": "3ad8879280f444fd143e6ade461d3ce0b8c69f504032c5f45c07aad991922b3f",
    "wasm-chartPie": "e00f3260266a0ed36829ad53453ee5c93def1db21d6f7b709b35529311eea95f",
    "wasm-chartBarOrder": "2d6951b0a5694d2ef6c0eb3de9fc326bb1a3bdfa75f0b41d76a93255bf397478",
    "wasm-chartLegendOrder": "572fa0e6be3b6c4164b89686d0cc51c57e2f12f1bed707582d61b05adbf39ad7",
    "wasm-chartStacked": "8344582ac41b68753e12d92ebb0dbd8835eef77c343e077d7d7e7e719da4fa0d",
    "wasm-chartArea": "59dcc94a6749bee320c16ba057b53dfdc22877842a7a26659b9d3b42b6f0b7f7",
    "wasm-chartPlotBand": "c0da1e12abe446d4dbba553db5f6ea150577d0874a320ee8cd12af9294a0472b",
    "wasm-chartCombo": "3056e0252c1c1df78500c439197c116fb427a2524d05bd6a110d823e63d5047a",
    "wasm-chartComboAreaPieLabels": "b53b38a367a98cc308266cd629d10c39e7bf38c73698293be3f0f6fb1a529899",
    "wasm-chartAreaOrder": "35a219c75dc618f4f8487e79d6da5af60e27fa4eb2fabbe1390f8e7cf27c0908",
    "wasm-chartPointColors": "39085a08e5392e5169d36ce3881d16d3ffaf7c691b6f65e06b628738ddc91466",
    "wasm-chartScatterAxes": "ba9fb83204341367085e67753d241f8b92f1d7cc96eda7f1864baecb8611a0ac",
    "wasm-chartTopRightLegend": "a0a36fbafbdcb2f0053817f106b3ec4d6c7a27a80841c96b1c3f1fbfb87122eb",
    "wasm-chartValueAxis": "47269f9038df4d9f4cb4f6f20d46985ad80a38239065d9c840f6dad7ab56b2a4",
    "wasm-chartLabelPositions": "8c7b6d2ae469813c97e619afff6763229c0c5cdec13e3a1522fc31d668419f45",
    "wasm-bulletType": "ce28b4623e9392fb8c8b9d01aa030c90ea8dc35f38be9e9b696834d0736de382",
    "wasm-chartFixedBound": "ff205f2fe4c6ce147005f29446706ac9f2e52e50514c96c769db2d59e9403019",
    "wasm-chartBar3DAxis": "25c89cf67b10d5ee733c69270bfa8edf777af90bb58c977a785f7ac6ece4d0f2",
    "wasm-chartOptionLabels": "45640b0a32243c90a360234ee035c0ab65c5a1e30e6b8843fef1a82f511e6acc",
    "wasm-chartOptionAxes": "ddf2c1868387cd1e6f4d9b1bf3b35ba356cd9ff1a85feb0354a5d0b69ea3327d",
    "wasm-chartCategoryOptions": "8541d3c97d75133277050b5b876104783d395080a409c0a0de1544da375b3da3",
    "wasm-chartDataTable": "dd8f0dec247576a4b157dfb89fbd03a574ad61a52760a4d8882ded8dc317d1d9",
    "wasm-chartPieBestFit": "6c7ae7e752e3a49c02ad652b1bd36a7d262d69f5882ac5cf5cc10c700cb20e20",
    "wasm-chartAutomaticScale": "1cf92ceb4460e5b99e032346907dd384ec3f5e35c753d79fd6d904c0f4f148fb",
    "wasm-chartDateAxis": "059eacdac2c87e96b9d2358243b2176e2b0bd5b867ac2c74ff4a8b41a473d5dd",
    "wasm-chartTypeOptions": "f2d68822bd540c5a2ed5c7b7f32847026da5e6168c501e9993e8d7f09947f7da",
    "wasm-chartOptionSources": "9d6476034a046a850a42ffd51ee56c1baa3f7f912a79d6910ce642a99c711d0d",
    "wasm-fallbackFaces": "3f0e840bfdd51ab6f39227bfdd1660becf4626c91bef0f978a4aac3ed7327727",
    "wasm-svg": "655268e09a87052e2f0f46d7b5fc5f6f8b7d436a7886fb4b189ad429c8cec24e",
    "wasm-gif": "4ab03748148d3ee110f2b7767497614ec513e227e8424ebe6b8fad499e2f82bf",
    "wasm-webp": "6aca049111e953889251e3c5bfef029395352fdf16f79741b837099eec1550be",
    "wasm-slides": "c05c3e7b74d770a288a68bba4025a14c863e4bd50df9b04117d479ecc75dc6f9",
    "wasm-chartLabels": "e3d21b5c3acdaee5cad516f6d8b102f973e23153e32c4c80f9df1f4c55545ee7",
    "wasm-components": "306e319217049fcad4206767ddeae2b4777b4986127957b685f076fcd603f576",
    "wasm-options": "910aac07c096ddcf0418bec42cccb5fd28f85b7cf5795d3d934103a2635eacfe",
    "wasm-tableOptions": "ec29f95067c21e9bd350d6feee9112a9fbac05f046c29da1bcfddff285a0c2db",
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
