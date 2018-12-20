# msoffice-pptx-rs

A library to deserialize pptx files in Rust.

## Overview
msoffice-pptx-rs is a low level deserializer for Microsoft's OfficeOpen XML pptx file format. It's still WIP, so expect API breaking changes.
Expect a more detailed documentation later.

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
