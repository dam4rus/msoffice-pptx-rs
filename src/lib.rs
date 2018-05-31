//extern crate quick_xml;

#[macro_use]
mod macros;
pub mod docprops;
pub mod relationship;
pub mod pml;
pub mod drawingml;
pub mod document;


#[cfg(test)]
mod tests {
    use drawingml;

    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
        let flip_mode = drawingml::TileFlipMode::from_string("none");
        if let drawingml::TileFlipMode::None = flip_mode {
            println!("TileFlipMode is none");
        }
    }
}
