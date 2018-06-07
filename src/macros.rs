macro_rules! decl_simple_type_enum {
    (pub enum $name:ident {
        $($variant:ident = $str_value:expr),*,
    }) => {
        pub enum $name {
            $($variant),*,
        }

        impl $name {
            pub fn from_string(s: &str) -> Result<$name, &'static str> {
                match s {
                    $($str_value => Ok($name::$variant)),*,
                    _ => Err("Cannot convert string to enum type $name"),
                }
            }
        }
    };
}