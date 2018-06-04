//extern crate quick_xml;

#[macro_use]
mod macros;
pub mod docprops;
pub mod relationship;
pub mod pml;
pub mod drawingml;
pub mod document;

// fn do_something(i: i32) -> Result<(), &'static str> {
//     match i {
//         1 => Ok(()),
//         _ => Err("error"),
//     }
// }

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

        //let i = do_something(1);
        //let i2 = do_something(2);

        // match i {
        //     Ok(()) => println!("ok"),
        //     Err(ref err) => println!("{}", err),
        // };

        // match i2 {
        //     Ok(()) => println!("ok"),
        //     Err(ref err) => println!("{}", err),
        // };
    }
}
