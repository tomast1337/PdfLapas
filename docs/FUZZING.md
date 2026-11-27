# Fuzzing the WASM renderer

The WASM renderer must not panic on any input. In the wasm build a panic traps the instance, and every
later render on the same `WasmRenderer` fails ([GHSA-fmx3-qwqc-h37c](https://github.com/tomast1337/PdfLapas/security/advisories/GHSA-fmx3-qwqc-h37c)).
`crates/fuzz` holds two [cargo-fuzz](https://github.com/rust-fuzz/cargo-fuzz) targets for `pdf-core`:

| Target        | Input                | What it does                                                                                                            |
| ------------- | -------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| `render_html` | mutated profile HTML | `Renderer::render_html` on one renderer shared by every input, as a server shares one `WasmRenderer`                    |
| `add_font`    | mutated font bytes   | `Renderer::add_font`, then a text deck (`fuzz_targets/add_font.html`) drawn in that face: shaping, outlines, subsetting |

A returned `Err` is the expected answer to invalid input. A panic, an abort, a timeout or running out of
memory is a bug.

## Requirements

cargo-fuzz needs a nightly toolchain (libFuzzer and sanitizer flags). The repository pins stable in
`rust-toolchain.toml` for the wasm build, so run the fuzzer with `+nightly`. The fuzz crate is its own
workspace and is not part of the wasm build or of the required CI checks. The nightly job (below) pins the
nightly date and the cargo-fuzz version in `.github/workflows/fuzz.yml`; use the same ones to reproduce what it
found.

```sh
rustup toolchain install nightly
cargo install cargo-fuzz --version 0.13.2 --locked
```

## Seeds

`render_html` starts from the `toHtml()` output of the shared visual fixture and of every WASM visual
deck, one document per slide:

```sh
bun run scripts/export-fuzz-seeds.ts   # writes crates/fuzz/corpus/render_html/
```

`add_font` starts from small real fonts: a TrueType (glyf), a CFF and a colour bitmap (CBDT) face. Subset
them to keep the inputs small, for example with fontTools:

```sh
mkdir -p crates/fuzz/corpus/add_font
pyftsubset LiberationSans-Regular.ttf --unicodes=U+0020-007E --output-file=crates/fuzz/corpus/add_font/glyf.ttf
pyftsubset NimbusSansNarrow-Bold.otf --unicodes=U+0020-007E --output-file=crates/fuzz/corpus/add_font/cff.otf
pyftsubset NotoColorEmoji.ttf --unicodes=U+1F600 --output-file=crates/fuzz/corpus/add_font/cbdt.ttf
```

The corpus and crash artifacts are not committed (`crates/fuzz/.gitignore`).

## Running

From `crates/`. `FUZZ_FONTS` lists the faces `render_html` registers, separated by colons. Without a font,
no glyph is drawn and the text paths are not reached. `fuzz/render_html.dict` holds the tags, attributes
and extreme numbers the mutator splices in.

```sh
cd crates
FUZZ_FONTS=/usr/share/fonts/liberation/LiberationSans-Regular.ttf \
  cargo +nightly fuzz run render_html -- -dict=fuzz/render_html.dict -max_total_time=900 -max_len=131072 -timeout=10 -rss_limit_mb=4096
cargo +nightly fuzz run add_font -- -max_total_time=900 -max_len=65536 -timeout=10 -rss_limit_mb=4096
```

A crash leaves its input in `crates/fuzz/artifacts/<target>/`. Reproduce it with
`cargo +nightly fuzz run <target> <artifact>`. Before fixing the crash, turn the input into a unit test in
`pdf-core`: the smallest HTML or font that still fails, with the error or warning it must give instead.
`cargo +nightly fuzz tmin <target> <artifact>` shrinks the input.

## Nightly job

`.github/workflows/fuzz.yml` runs every night, and on demand from the Actions tab (`workflow_dispatch`, with the
seconds per target and the property scale as inputs). It is not a required check.

- **Rust targets.** `render_html` and `add_font` run for 600 seconds each, with the seeds and limits above.
  The corpus each night grows is kept in the Actions cache, so the next night starts from it.
- **TypeScript properties.** The fast-check properties in `tests/*.fuzz.test.ts` (PptxGenJS API options,
  `parseModel()`, the image pixel scan, `PdfkitRenderer.render()` over mutated HTML) run with a new seed and 20
  times their usual runs. On every `bun test` they use a fixed seed (`tests/fuzzConfig.ts`).

When a property fails, the log prints its seed. Reproduce it locally:

```sh
FUZZ_SEED=<seed> FUZZ_SCALE=20 bun test --timeout 1800000 tests/<file>.fuzz.test.ts
```

When a Rust target fails, its crash input is uploaded as the `fuzz-crash-<target>` artifact and kept for 7 days.
Anyone who can read the repository can download it. Triage it within that week: if it is exploitable, report it as
a private security advisory and do not discuss it in a public issue. Failures are never filed automatically.
