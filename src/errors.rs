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