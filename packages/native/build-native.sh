#!/usr/bin/env sh
# Build the native renderer (crates/pdf-ffi) into its platform package, packages/native-<target>/pdflapas_native.so,
# where src/bun.ts loads it from (the workspace links every platform package, as an install links the matching one).
#   bun run build:native                    this host's glibc target (linux-x64-gnu or linux-arm64-gnu)
#   bun run build:native linux-x64-musl     a musl target, in the pinned Alpine Rust image (RUST_MUSL_IMAGE)
# No cross-compiling: the target's CPU must be this host's, so CI builds each target on a runner of its CPU.
# glibc targets use the toolchain from rust-toolchain.toml, through scripts/with-rust.sh.
set -eu
cd "$(dirname "$0")/../.."
. crates/wasm-toolchain.env
case "$(uname -sm)" in
    "Linux x86_64") arch=x64 ;;
    "Linux aarch64") arch=arm64 ;;
    *)
        echo "build:native builds linux x64 and arm64 only (darwin and windows are later, #612); this host is $(uname -sm)" >&2
        exit 1
        ;;
esac
target=${1:-linux-$arch-gnu}
case "$target" in
    "linux-$arch-gnu" | "linux-$arch-musl") ;;
    *)
        echo "build:native cannot build $target on this linux-$arch host; use linux-$arch-gnu or linux-$arch-musl" >&2
        exit 1
        ;;
esac
root=$PWD
cargo_build="cargo build --manifest-path crates/Cargo.toml -p pdf-ffi --profile native --locked"
case "$target" in
    *-gnu)
        # Panic messages embed source paths: remap them as packages/wasm/build-wasm.sh does.
        RUSTFLAGS="--remap-path-prefix=$root=/pdflapas --remap-path-prefix=${CARGO_HOME:-$HOME/.cargo}=/cargo"
        export RUSTFLAGS
        sh scripts/with-rust.sh $cargo_build
        built=crates/target/native/libpdflapas_native.so
        ;;
    *-musl)
        # Alpine's host target is musl. Rust links musl statically by default, and a cdylib cannot be: -crt-static
        # links musl's libc.so, which Bun on Alpine has loaded already. Its own target dir keeps the glibc build's cache.
        # RUSTUP_TOOLCHAIN names the pinned compiler, so an image without it fails instead of building with another.
        mkdir -p crates/target
        docker run --rm --user "$(id -u):$(id -g)" --volume "$root:/src" --workdir /src \
            --env HOME=/tmp --env CARGO_HOME=/src/crates/target/docker-cargo-home --env CARGO_TARGET_DIR=/src/crates/target/musl \
            --env RUSTUP_TOOLCHAIN="$(sed -n 's/^channel = "\(.*\)"$/\1/p' rust-toolchain.toml)" \
            --env RUSTFLAGS="-C target-feature=-crt-static --remap-path-prefix=/src=/pdflapas --remap-path-prefix=/src/crates/target/docker-cargo-home=/cargo" \
            "$RUST_MUSL_IMAGE" $cargo_build
        built=crates/target/musl/native/libpdflapas_native.so
        ;;
esac
cp "$built" "packages/native-$target/pdflapas_native.so"
echo "packages/native-$target/pdflapas_native.so $(wc -c < "packages/native-$target/pdflapas_native.so" | tr -d ' ') B"
