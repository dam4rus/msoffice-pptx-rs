use std::fmt;
use std::error::Error;

#[derive(Debug)]
pub struct MissingAttributeError {
    pub attr: &'static str,
}

impl MissingAttributeError {
    pub fn new(attr: &'static str) -> Self {
        Self {
            attr,
        }
    }
}

impl fmt::Display for MissingAttributeError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Missing required attribute: {}", self.attr)
    }
}

impl Error for MissingAttributeError {
    fn description(&self) -> &str {
        "Missing required attribute"
    }
}

/// Error indicating that an xml element doesn't have a required child node
#[derive(Debug)]
pub struct MissingChildNodeError {
    pub child_node: &'static str,
}

impl MissingChildNodeError {
    pub fn new(child_node: &'static str) -> Self {
        Self {
            child_node
        }
    }
}

impl fmt::Display for MissingChildNodeError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Xml element is missing a required child element: {}", self.child_node)
    }
}

impl Error for MissingChildNodeError {
    fn description(&self) -> &str {
        "Xml element missing required child element"
    }
}

/// Error indicating that an xml element is not a member of a given element group
#[derive(Debug)]
pub struct NotGroupMemberError {
    pub group: &'static str,
}

impl fmt::Display for NotGroupMemberError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Xml element is not a member of {} group", self.group)
    }
}

impl Error for NotGroupMemberError {
    fn description(&self) -> &str {
        "Xml element is not a group member error"
    }
}

/// Chained error type for all possible xml error
#[derive(Debug)]
pub enum XmlError {
    Attribute(MissingAttributeError),
    ChildNode(MissingChildNodeError),
    NotGroupMember(NotGroupMemberError),
}

impl fmt::Display for XmlError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            XmlError::Attribute(ref err) => err.fmt(f),
            XmlError::ChildNode(ref err) => err.fmt(f),
            XmlError::NotGroupMember(ref err) => err.fmt(f),
        }
    }
}

impl Error for XmlError {
    fn description(&self) -> &str {
        "Xml element missing required attribute or child node"
    }
}

impl From<MissingAttributeError> for XmlError {
    fn from(error: MissingAttributeError) -> Self {
        XmlError::Attribute(error)
    }
}

impl From<MissingChildNodeError> for XmlError {
    fn from(error: MissingChildNodeError) -> Self {
        XmlError::ChildNode(error)
    }
}

impl From<NotGroupMemberError> for XmlError {
    fn from(error: NotGroupMemberError) -> Self {
        XmlError::NotGroupMember(error)
    }
}

/// Error indicating that an xml element's attribute is not a valid bool value
/// Valid bool values are: true, false, 0, 1
#[derive(Debug)]
pub struct ParseBoolError<'a> {
    pub attr_value: &'a str,
}

impl<'a> fmt::Display for ParseBoolError<'a> {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Xml attribute is not a valid bool value: {}", self.attr_value)
    }
}

impl<'a> Error for ParseBoolError<'a> {
    fn description(&self) -> &str {
        "Xml attribute is not a valid bool value"
    }
}