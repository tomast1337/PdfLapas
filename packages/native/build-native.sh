#!/usr/bin/env sh
# Build the native renderer (crates/pdf-ffi) for this host and copy it next to its loader in packages/native/src/, named
# for the platform the way src/bun.ts looks it up: pdflapas_native.<process.platform>-<process.arch>.<bun:ffi suffix>.
# Toolchain from rust-toolchain.toml, through scripts/with-rust.sh (the pinned Docker image builds linux-x64).
# shortcut: linux-x64 only, built by hand and kept out of `bun run build` and CI; per-platform packages are phase 2 (#612).
set -eu
cd "$(dirname "$0")/../.."
if [ "$(uname -sm)" != "Linux x86_64" ]; then
    echo "build:native is a linux-x64 prototype; this host is $(uname -sm)" >&2
    exit 1
fi
root=$PWD
# Panic messages embed source paths: remap them as packages/wasm/build-wasm.sh does.
RUSTFLAGS="--remap-path-prefix=$root=/pdflapas --remap-path-prefix=${CARGO_HOME:-$HOME/.cargo}=/cargo"
export RUSTFLAGS
sh scripts/with-rust.sh cargo build --manifest-path crates/Cargo.toml -p pdf-ffi --profile native --locked
cp crates/target/native/libpdflapas_native.so packages/native/src/pdflapas_native.linux-x64.so
echo "packages/native/src/pdflapas_native.linux-x64.so $(wc -c < packages/native/src/pdflapas_native.linux-x64.so | tr -d ' ') B"
