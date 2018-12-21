use std::error::Error;
use std::fmt::{Display, Formatter, Result};

/// An error indicating that an xml element doesn't have an attribute that's marked as required in the schema
#[derive(Debug, Clone)]
pub struct MissingAttributeError {
    pub node_name: String,
    pub attr: &'static str,
}

impl MissingAttributeError {
    pub fn new<T>(node_name: T, attr: &'static str) -> Self
    where
        T: Into<String>,
    {
        Self {
            node_name: node_name.into(),
            attr,
        }
    }
}

impl Display for MissingAttributeError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "Xml element '{}' is missing a required attribute: {}",
            self.node_name, self.attr
        )
    }
}

impl Error for MissingAttributeError {
    fn description(&self) -> &str {
        "Missing required attribute"
    }
}

/// An error indicating that an xml element doesn't have a child node that's marked as required in the schema
#[derive(Debug, Clone)]
pub struct MissingChildNodeError {
    pub node_name: String,
    pub child_node: &'static str,
}

impl MissingChildNodeError {
    pub fn new<T>(node_name: T, child_node: &'static str) -> Self
    where
        T: Into<String>,
    {
        Self {
            node_name: node_name.into(),
            child_node,
        }
    }
}

impl Display for MissingChildNodeError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "Xml element '{}' is missing a required child element: {}",
            self.node_name, self.child_node
        )
    }
}

impl Error for MissingChildNodeError {
    fn description(&self) -> &str {
        "Xml element missing required child element"
    }
}

/// An error indicating that an xml element is not a member of a given element group
#[derive(Debug, Clone)]
pub struct NotGroupMemberError {
    node_name: String,
    group: &'static str,
}

impl NotGroupMemberError {
    pub fn new<T>(node_name: T, group: &'static str) -> Self
    where
        T: Into<String>,
    {
        Self {
            node_name: node_name.into(),
            group,
        }
    }
}

impl Display for NotGroupMemberError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "XmlNode '{}' is not a member of {} group",
            self.node_name, self.group
        )
    }
}

impl Error for NotGroupMemberError {
    fn description(&self) -> &str {
        "Xml element is not a group member error"
    }
}

#[derive(Debug, Clone, Copy)]
pub enum Limit {
    Value(u32),
    Unbounded,
}

impl Display for Limit {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        match self {
            Limit::Value(val) => write!(f, "{}", val),
            Limit::Unbounded => write!(f, "unbounded"),
        }
    }
}

/// An error indicating that the xml element violates either minOccurs or maxOccurs of the schema
#[derive(Debug, Clone)]
pub struct LimitViolationError {
    node_name: String,
    violating_node_name: &'static str,
    min_occurs: Limit,
    max_occurs: Limit,
    occurs: u32,
}

impl LimitViolationError {
    pub fn new<T>(
        node_name: T,
        violating_node_name: &'static str,
        min_occurs: Limit,
        max_occurs: Limit,
        occurs: u32,
    ) -> Self
    where
        T: Into<String>,
    {
        LimitViolationError {
            node_name: node_name.into(),
            violating_node_name,
            min_occurs,
            max_occurs,
            occurs,
        }
    }
}

impl Display for LimitViolationError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "Element {} violates the limits of occurance in element: {}. minOccurs: {}, maxOccurs: {}, occurance: {}",
            self.node_name, self.violating_node_name, self.min_occurs, self.max_occurs, self.occurs,
        )
    }
}

impl Error for LimitViolationError {
    fn description(&self) -> &str {
        "Occurance limit violation"
    }
}

/// Chained error type for all possible xml error
#[derive(Debug, Clone)]
pub enum XmlError {
    Attribute(MissingAttributeError),
    ChildNode(MissingChildNodeError),
    NotGroupMember(NotGroupMemberError),
    LimitViolation(LimitViolationError),
}

impl Display for XmlError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
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

/// An error indicating that the parsed xml document is invalid
#[derive(Debug, Clone, Copy)]
pub struct InvalidXmlError {}

impl InvalidXmlError {
    pub fn new() -> Self {
        Self {}
    }
}

impl Display for InvalidXmlError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
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
#[derive(Debug, Clone)]
pub struct ParseBoolError {
    pub attr_value: String,
}

impl ParseBoolError {
    pub fn new<T>(attr_value: T) -> Self
    where
        T: Into<String>,
    {
        Self {
            attr_value: attr_value.into(),
        }
    }
}

impl Display for ParseBoolError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "Xml attribute is not a valid bool value: {}", self.attr_value)
    }
}

impl Error for ParseBoolError {
    fn description(&self) -> &str {
        "Xml attribute is not a valid bool value"
    }
}

/// Error indicating that a string cannot be converted to an enum type
#[derive(Debug, Clone, Copy)]
pub struct ParseEnumError {
    enum_name: &'static str,
}

impl ParseEnumError {
    pub fn new(enum_name: &'static str) -> Self {
        Self { enum_name }
    }
}

impl Display for ParseEnumError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "Cannot convert string to {}", self.enum_name)
    }
}

impl Error for ParseEnumError {
    fn description(&self) -> &str {
        "Cannot convert string to enum"
    }
}

#[derive(Debug, Clone, Copy)]
/// Error indicating that parsing an AdjCoordinate or AdjAngle has failed
pub struct AdjustParseError {}

impl Display for AdjustParseError {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "AdjCoordinate or AdjAngle parse error")
    }
}

impl Error for AdjustParseError {
    fn description(&self) -> &str {
        "Adjust parse error"
    }
}
