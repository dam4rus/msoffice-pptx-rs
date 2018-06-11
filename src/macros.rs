macro_rules! decl_simple_type_enum {
    (pub enum $name:ident {
        $($variant:ident = $str_value:expr),*,
    }) => {
        pub enum $name {
            $($variant),*,
        }

        impl $name {
            pub fn from_string(s: &str) -> Result<$name, String> {
                match s {
                    $($str_value => Ok($name::$variant)),*,
                    _ => Err(format!("Cannot convert string to {}", stringify!($name))),
                }
            }
        }
    };
}