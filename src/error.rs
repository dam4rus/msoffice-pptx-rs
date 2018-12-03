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
    node_name: String,
    group: &'static str,
}

impl NotGroupMemberError {
    pub fn new(node_name: String, group: &'static str) -> Self {
        Self {
            node_name,
            group,
        }
    }
}

impl fmt::Display for NotGroupMemberError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "XmlNode '{}' is not a member of {} group", self.node_name, self.group)
    }
}

impl Error for NotGroupMemberError {
    fn description(&self) -> &str {
        "Xml element is not a group member error"
    }
}

#[derive(Debug)]
pub enum Limit {
    Value(u32),
    Unbounded,
}

impl fmt::Display for Limit {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            Limit::Value(val) => write!(f, "{}", val),
            Limit::Unbounded => write!(f, "unbounded"),
        }
    }
}

/// Error indicating that the element violates either minOccurs or maxOccurs
#[derive(Debug)]
pub struct LimitViolationError {
    element_name: &'static str,
    min_occurs: Limit,
    max_occurs: Limit,
    occurs: u32,
}

impl LimitViolationError {
    pub fn new(
        element_name: &'static str,
        min_occurs: Limit,
        max_occurs: Limit,
        occurs: u32
    ) -> Self {
        LimitViolationError {
            element_name,
            min_occurs,
            max_occurs,
            occurs,
        }
    }
}

impl fmt::Display for LimitViolationError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(
            f,
            "Element {} violates the limits of occurance. minOccurs: {}, maxOccurs: {}, occurance: {}",
            self.element_name,
            self.min_occurs,
            self.max_occurs,
            self.occurs
        )
    }
}

impl Error for LimitViolationError {
    fn description(&self) -> &str {
        "Occurance limit violation"
    }
}

/// Chained error type for all possible xml error
#[derive(Debug)]
pub enum XmlError {
    Attribute(MissingAttributeError),
    ChildNode(MissingChildNodeError),
    NotGroupMember(NotGroupMemberError),
    LimitViolation(LimitViolationError),
}

impl fmt::Display for XmlError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            XmlError::Attribute(ref err) => err.fmt(f),
            XmlError::ChildNode(ref err) => err.fmt(f),
            XmlError::NotGroupMember(ref err) => err.fmt(f),
            XmlError::LimitViolation(ref err) => err.fmt(f),
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

impl From<LimitViolationError> for XmlError {
    fn from(error: LimitViolationError) -> Self {
        XmlError::LimitViolation(error)
    }
}

/// Error indicating that the parsed xml document is invalid
#[derive(Debug)]
pub struct InvalidXmlError {
}

impl InvalidXmlError {
    pub fn new() -> Self {
        Self {}
    }
}

impl fmt::Display for InvalidXmlError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Invalid xml document")
    }
}

impl Error for InvalidXmlError {
    fn description(&self) -> &str {
        "Invalid xml error"
    }
}


/// Error indicating that an xml element's attribute is not a valid bool value
/// Valid bool values are: true, false, 0, 1
#[derive(Debug)]
pub struct ParseBoolError {
    pub attr_value: String,
}

impl ParseBoolError {
    pub fn new(attr_value: String) -> Self {
        Self {
            attr_value,
        }
    }
}

impl fmt::Display for ParseBoolError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Xml attribute is not a valid bool value: {}", self.attr_value)
    }
}

impl Error for ParseBoolError {
    fn description(&self) -> &str {
        "Xml attribute is not a valid bool value"
    }
}

/// Error indicating that a string cannot be converted to an enum type
pub struct ParseEnumError {
    enum_name: &'static str,
}

impl ParseEnumError {
    pub fn new(enum_name: &'static str) -> Self {
        Self {
            enum_name,
        }
    }
}

impl fmt::Display for ParseEnumError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Cannot convert string to {}", self.enum_name)
    }
}

impl Error for ParseEnumError {
    fn description(&self) -> &str {
        "Cannot convert string to enum"
    }
}