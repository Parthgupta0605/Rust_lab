use std::cell::RefCell;
use std::rc::Rc;

// Maximum input length (matches #define MAX_INPUT_LEN 35)
pub const MAX_INPUT_LEN: usize = 35;

// Forward declarations for AVL and Stack structures
// Define them in separate modules/files as per your architecture
use crate::avl::{AvlNode, Link as AvlLink}; // Assuming `avl.rs` defines AVL tree
use crate::stack::{StackNode, StackLink};   // Assuming `stack.rs` defines Stack

// Cell structure equivalent
#[derive(Clone)]
pub struct Cell {
    pub val: i32,                           // Value of the cell
    pub expression: String,                 // Expression stored as a String
    pub status: i32,                        // Status to determine if it has ERR
    pub dependencies: AvlLink,              // AVL tree of dependencies
    pub dependents: StackLink,              // Stack of dependents
}

impl Cell {
    pub fn new(val: i32, expression: &str, status: i32) -> Rc<RefCell<Self>> {
        Rc::new(RefCell::new(Self {
            val,
            expression: expression.chars().take(MAX_INPUT_LEN).collect(),
            status,
            dependencies: None,
            dependents: None,
        }))
    }
}
