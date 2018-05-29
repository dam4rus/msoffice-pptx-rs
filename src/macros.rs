macro_rules! decl_oox_enum {
    (pub enum $name:ident {
        $($variant:ident = $str_value:expr),*,
    }) => {
        pub enum $name {
            $($variant),*,
        }

        impl $name {
            pub fn from_string(s: &String) -> $name {
                match s.as_str() {
                    $($str_value => $name::$variant),*,
                    _ => panic!("Cannot convert string to enum type $name"),
                }
            }
        }
    };
}