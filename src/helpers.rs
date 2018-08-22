use quick_xml;
use std::error::Error;
use std::fmt;
use std::str;

/// XmlAttributeParseError
#[derive(Debug)]
pub enum XmlAttributeParseError<T>
where
    T: str::FromStr,
{
    Utf8Error(str::Utf8Error),
    ParseError(T::Err),
}

impl<T> fmt::Display for XmlAttributeParseError<T>
where
    T: str::FromStr,
    T::Err: fmt::Display,
{
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        use self::XmlAttributeParseError::*;
        match *self {
            Utf8Error(ref err) => err.fmt(f),
            ParseError(ref err) => err.fmt(f),
        }
    }
}

impl<T> Error for XmlAttributeParseError<T>
where
    T: str::FromStr + fmt::Debug,
    T::Err: fmt::Display + fmt::Debug,
{
    fn description(&self) -> &str {
        "Xml attribute parse error"
    }
}

/// SimpleElementParseError
#[derive(Debug)]
pub enum SimpleElementParseError<T>
where
    T: str::FromStr,
    T::Err: fmt::Debug,
{
    MissingAttribute,
    ParseError(XmlAttributeParseError<T>),
}

impl<T> fmt::Display for SimpleElementParseError<T>
where
    T: str::FromStr,
    T::Err: fmt::Debug + fmt::Display,
{
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        use self::SimpleElementParseError::*;
        match *self {
            MissingAttribute => write!(f, "Xml element doesn't have the required attribute"),
            ParseError(ref err) => err.fmt(f),
        }
    }
}

impl<T> Error for SimpleElementParseError<T>
where
    T: str::FromStr + fmt::Debug + fmt::Display,
    T::Err: fmt::Debug + fmt::Display,
{
    fn description(&self) -> &str {
        "Simple element parse error"
    }
}

/// Parse an xml attribute and return a result. This should be used to parse required attributes
pub fn parse_xml_attribute<T>(attr: &[u8]) -> Result<T, XmlAttributeParseError<T>>
where
    T: str::FromStr,
{
    let attr_str = match str::from_utf8(&attr) {
        Ok(s) => s,
        Err(err) => return Err(XmlAttributeParseError::Utf8Error(err)),
    };

    match attr_str.parse::<T>() {
        Ok(value) => Ok(value),
        Err(err) => Err(XmlAttributeParseError::ParseError(err)),
    }
}

/// Parse an xml attribute. On success returns the parsed value, on failure returns the provided default value
pub fn parse_optional_xml_attribute<T>(attr: &[u8], default: T) -> T
where
    T: str::FromStr,
{
    let attr_str = match str::from_utf8(&attr) {
        Ok(s) => s,
        Err(_) => return default,
    };

    match attr_str.parse::<T>() {
        Ok(value) => value,
        Err(_) => default,
    }
}

/// Parse a single attribute of an xml element
pub fn parse_xml_element_attribute<T>(
    xml_element: &quick_xml::events::BytesStart,
    attr_name: &[u8],
) -> Result<T, SimpleElementParseError<T>>
where
    T: str::FromStr,
    T::Err: fmt::Debug,
{
    if let Some(attr) = xml_element.attributes().next() {
        if let Ok(a) = attr {
            if a.key == attr_name {
                return match parse_xml_attribute(&a.value) {
                    Ok(val) => Ok(val),
                    Err(err) => Err(SimpleElementParseError::ParseError(err)),
                };
            }
        }
    }

    Err(SimpleElementParseError::MissingAttribute)
}

/// Parses a single attribute of an xml element, returning None if the element has no such attribute
pub fn parse_optional_xml_element_attribute<T>(
    xml_element: &quick_xml::events::BytesStart,
    attr_name: &[u8],
    default: T,
) -> Option<T>
where
    T: str::FromStr,
{
    if let Some(attr) = xml_element.attributes().next() {
        if let Ok(a) = attr {
            if a.key == attr_name {
                return Some(parse_optional_xml_attribute(&a.value, default));
            }
        }
    }

    None
}

/// Iterate through child elements of an xml element, stoping if `callback` returns true
pub fn iterate_xml_element_childs<F>(
    xml_element: &quick_xml::events::BytesStart,
    xml_reader: &mut quick_xml::Reader<&[u8]>,
    mut callback: F,
) where
    F: FnMut(&quick_xml::events::BytesStart, &mut quick_xml::Reader<&[u8]>) -> bool,
{
    let mut buffer = Vec::new();
    loop {
        use quick_xml::events::Event;
        match xml_reader.read_event(&mut buffer) {
            Ok(Event::Start(ref element)) => {
                if callback(element, xml_reader) {
                    break;
                }
            }
            Ok(Event::End(ref element)) => {
                if element.local_name() == xml_element.local_name() {
                    break;
                }
            }
            _ => (),
        }
    }
}
