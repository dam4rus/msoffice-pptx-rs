use std::collections::HashMap;
use quick_xml::{Reader};
use quick_xml::events::{BytesStart};

/// Represents an implementation independent xml node
pub struct XmlNode {
    pub name: String,
    pub child_nodes: Vec<XmlNode>,
    pub attributes: HashMap<String, String>,
}

impl XmlNode {
    pub fn new(name: &str) -> XmlNode {
        XmlNode {
            name: String::from(name),
            child_nodes: Vec::new(),
            attributes: HashMap::new(),
        }
    }

    pub fn from_str(xml_string: &str) -> Option<XmlNode> {
        let mut xml_reader = quick_xml::Reader::from_str(xml_string);
        let mut root_node = None;
        let mut buffer = Vec::new();
        loop {
            match xml_reader.read_event(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element)) => {
                    root_node = XmlNode::from_quick_xml_element(element);
                    if let Some(ref mut node) = root_node {
                        node.child_nodes = XmlNode::parse_child_elements(element, &mut xml_reader);
                    }
                },
                Ok(quick_xml::events::Event::Eof) => break,
                _ => (),
            }

            buffer.clear();
        }

        root_node
    }

    pub fn local_name(&self) -> &str {
        match self.name.find(':') {
            Some(idx) => self.name.split_at(idx + 1).1,
            None => self.name.as_str(),
        }
    }

    pub fn attribute(&self, attr_name: &str) -> Option<&String> {
        self.attributes.get(attr_name)
        //self.attributes[attr_name].as_str()
    }

    fn from_quick_xml_element(xml_element: &quick_xml::events::BytesStart) -> Option<XmlNode> {
        let name_str = match std::str::from_utf8(xml_element.name()) {
            Ok(s) => s,
            Err(_) => return None,
        };

        let mut node = XmlNode::new(name_str);

        for attr in xml_element.attributes() {
            if let Ok(a) = attr {
                let key_str = match std::str::from_utf8(&a.key) {
                    Ok(s) => s,
                    Err(_) => return None,
                };

                let value_str = match std::str::from_utf8(&a.value) {
                    Ok(s) => s,
                    Err(_) => return None,
                };

                node.attributes.insert(String::from(key_str), String::from(value_str));
            }
        }

        Some(node)
    }

    fn parse_child_elements(
        xml_element: &BytesStart,
        xml_reader: &mut Reader<&[u8]>,
    ) -> Vec<XmlNode> {
        let mut child_nodes = Vec::new();

        let mut buffer = Vec::new();
        loop {
            match xml_reader.read_event(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element)) => {
                    if let Some(mut node) = XmlNode::from_quick_xml_element(element) {
                        node.child_nodes = XmlNode::parse_child_elements(element, xml_reader);
                        child_nodes.push(node);
                    }
                },
                Ok(quick_xml::events::Event::Empty(ref element)) => {
                    if let Some(mut node) = XmlNode::from_quick_xml_element(element) {
                        child_nodes.push(node);
                    }
                },
                Ok(quick_xml::events::Event::End(ref element)) => {
                    if element.name() == xml_element.name() {
                        break;
                    }
                }
                _ => (),
            }

            buffer.clear();
        }

        child_nodes
    }
}

/// Parse an xml attribute.
/// On success it returns `Some` with the parsed value, on failure it prints the error and returns `None`
pub fn parse_optional_xml_attribute<T>(attr: &str) -> Option<T>
where
    T: std::str::FromStr,
    T::Err: std::fmt::Debug + std::fmt::Display,
{
    match attr.parse::<T>() {
        Ok(value) => Some(value),
        Err(err) => {
            println!("{}", err);
            None
        }
    }
}

pub fn parse_xml_bool(value: &str) -> Result<bool, ::error::ParseBoolError> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(::error::ParseBoolError { attr_value: value }),
    }
}