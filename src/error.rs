use std::fmt;
use std::error::Error;

#[derive(Debug)]
pub struct MissingAttributeError {
    pub attr: &'static str,
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

/// Error indicating that an xml element doesn't have a required child node
#[derive(Debug)]
pub struct MissingChildNodeError {
    pub child_node: &'static str,
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