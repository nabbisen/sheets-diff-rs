# sheets-diff

Collect diff between office sheets written in Rust

[![crates.io](https://img.shields.io/crates/v/sheets-diff?label=latest)](https://crates.io/crates/sheets-diff)
[![Documentation](https://docs.rs/sheets-diff/badge.svg?version=latest)](https://docs.rs/sheets-diff/latest)
[![Dependency Status](https://deps.rs/crate/sheets-diff/latest/status.svg)](https://deps.rs/crate/sheets-diff/latest)
[![Releases Workflow](https://github.com/nabbisen/sheets-diff-rs/actions/workflows/release.yml/badge.svg)](https://github.com/nabbisen/sheets-diff-rs/actions/workflows/)
[![License](https://img.shields.io/github/license/nabbisen/sheets-diff-rs)](https://github.com/nabbisen/sheets-diff-rs/blob/main/LICENSE)

## Features

With `.xlsx`, Microsoft Office Excel:

- Get unified diff between two files
- Get serde-ready diff
    - Note: `serde` feature is required: `cargo add sheets-diff -F serde`

## Simple run

```console
$ # via executable available in Releases
$ ./sheets-diff <file1> <file2>

$ # via cargo
$ # first `cargo add sheets-diff`
$ cargo run -- <file1> <file2>
```

### Output example

```console
--- ./file1.xlsx (sheet names)
+++ ./file2.xlsx (sheet names)
- RemovedSheet
+ AddedSheet
--- ./file1.xlsx [Sheet1]
+++ ./file2.xlsx [Sheet1]
@@ A1(1,1) value @@
- 1
@@ D10(10,4) formula @@
- 1+1
+ 2*4
```

## Acknowledgements

Depends on:

- [tafia](https://github.com/tafia)'s [calamine](https://github.com/tafia/calamine) and [quick-xml](https://github.com/tafia/quick-xml)
- Also big thanks to [zip-rs/zip2](https://github.com/zip-rs/zip2) etc.
