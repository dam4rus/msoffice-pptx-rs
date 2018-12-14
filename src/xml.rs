use std::collections::HashMap;
use quick_xml::{Reader};
use quick_xml::events::{BytesStart, Event};
use ::error::InvalidXmlError;

/// Represents an implementation independent xml node
#[derive(Debug)]
pub struct XmlNode {
    pub name: String,
    pub child_nodes: Vec<XmlNode>,
    pub attributes: HashMap<String, String>,
    pub text: Option<String>,
}

impl ::std::fmt::Display for XmlNode {
    fn fmt(&self, f: &mut ::std::fmt::Formatter) -> ::std::fmt::Result {
        write!(f, "name: {}", self.name)
    }
}

impl XmlNode {
    pub fn new(name: &str) -> Self {
        Self {
            name: String::from(name),
            child_nodes: Vec::new(),
            attributes: HashMap::new(),
            text: None,
        }
    }

    pub fn from_str(xml_string: &str) -> Result<Self, InvalidXmlError> {
        let mut xml_reader = Reader::from_str(xml_string);
        let mut buffer = Vec::new();
        loop {
            match xml_reader.read_event(&mut buffer) {
                Ok(Event::Start(ref element)) => {
                    let mut root_node = Self::from_quick_xml_element(element).map_err(|_| InvalidXmlError::new())?;
                    root_node.child_nodes = Self::parse_child_elements(&mut root_node, element, &mut xml_reader).map_err(|_| InvalidXmlError::new())?;
                    return Ok(root_node);
                },
                Ok(Event::Eof) => break,
                _ => (),
            }

            buffer.clear();
        }

        Err(InvalidXmlError::new())
    }

    pub fn local_name(&self) -> &str {
        match self.name.find(':') {
            Some(idx) => self.name.split_at(idx + 1).1,
            None => self.name.as_str(),
        }
    }

    pub fn attribute(&self, attr_name: &str) -> Option<&String> {
        self.attributes.get(attr_name)
    }

    fn from_quick_xml_element(xml_element: &quick_xml::events::BytesStart) -> Result<Self, ::std::str::Utf8Error> {
        let name = ::std::str::from_utf8(xml_element.name())?;
        let mut node = Self::new(name);

        for attr in xml_element.attributes() {
            if let Ok(a) = attr {
                let key_str = ::std::str::from_utf8(&a.key)?;
                let value_str = ::std::str::from_utf8(&a.value)?;
                node.attributes.insert(String::from(key_str), String::from(value_str));
            }
        }

        Ok(node)
    }

    fn parse_child_elements(
        xml_node: &mut Self,
        xml_element: &BytesStart,
        xml_reader: &mut Reader<&[u8]>,
    ) -> Result<Vec<Self>, ::std::str::Utf8Error> {
        let mut child_nodes = Vec::new();

        let mut buffer = Vec::new();
        loop {
            match xml_reader.read_event(&mut buffer) {
                Ok(Event::Start(ref element)) => {
                    let mut node = Self::from_quick_xml_element(element)?;
                    node.child_nodes = Self::parse_child_elements(&mut node, element, xml_reader)?;
                    child_nodes.push(node);
                },
                Ok(Event::Text(text)) => {
                    xml_node.text = text.unescape_and_decode(xml_reader).ok();
                }
                Ok(Event::Empty(ref element)) => {
                    let node = Self::from_quick_xml_element(element)?;
                    child_nodes.push(node);
                },
                Ok(Event::End(ref element)) => {
                    if element.name() == xml_element.name() {
                        break;
                    }
                }
                Ok(Event::Eof) => {
                    break;
                }
                _ => (),
            }

            buffer.clear();
        }

        Ok(child_nodes)
    }
}

pub fn parse_xml_bool(value: &str) -> Result<bool, ::error::ParseBoolError> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(::error::ParseBoolError::new(String::from(value))),
    }
}