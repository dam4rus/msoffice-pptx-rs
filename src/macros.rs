macro_rules! decl_simple_type_enum {
    (pub enum $name:ident {
        $($variant:ident = $str_value:expr),*,
    }) => {
        #[derive(Debug)]
        pub enum $name {
            $($variant),*,
        }

        impl ::std::str::FromStr for $name {
            type Err = String;

            fn from_str(s: &str) -> Result<Self, Self::Err> {
                match s {
                    $($str_value => Ok($name::$variant)),*,
                    _ => Err(format!("Cannot convert string to {}", stringify!($name))),
                }
            }
        }
    };
}