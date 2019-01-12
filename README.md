# msoffice-pptx-rs

A library to deserialize pptx files in Rust.

[![Latest version](https://img.shields.io/crates/v/msoffice_pptx.svg)](https://crates.io/crates/msoffice_pptx)
[![Documentation](https://docs.rs/msoffice_pptx/badge.svg)](https://docs.rs/msoffice_pptx)

## Overview

msoffice-pptx-rs is a low level deserializer for Microsoft's OfficeOpen XML pptx file format. It's still WIP, so expect API breaking changes.

The Office Open XML file formats are described by the [ECMA-376 standard](http://www.ecma-international.org/publications/standards/Ecma-376.htm).
The types represented in this library are generated from the Transitional XML Schema's, which is described in
[ECMA-376 4th edition Part 4](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%204%20-%20Transitional%20Migration%20Features.zip), "pml.xsd" file.

Documentation is generated from the "Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf" file, found in [ECMA-376 4th edition Part 1](http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-376,%20Fourth%20Edition,%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.zip)

## Simple usage

```rust
extern crate msoffice_pptx;

use msoffice_pptx::document::PPTXDocument;

pub fn main() {
  let document = PPTXDocument::from_file(Path::new("test.pptx")).unwrap();
  
  for (slide_path, slide) in &document.slide_map {
    // Do something with slides
  }
}
```
