use std::cell::RefCell;
use std::rc::Rc;
pub const MAX_INPUT_LEN_CELL: usize = 35;
pub type CellRef = Rc<RefCell<Cell>>;
use crate::avl::Link as AvlLink; // Assuming `avl.rs` defines AVL tree
use crate::stack::StackLink; // Assuming `stack.rs` defines Stack

// Cell structure equivalent
#[derive(Clone)]
pub struct Cell {
    pub val: i32,              // Value of the cell
    pub expression: String,    // Expression stored as a String
    pub status: i32,           // Status to determine if it has ERR
    pub dependencies: AvlLink, // AVL tree of dependencies
    pub dependents: StackLink, // Stack of dependents
}

impl Cell {
    pub fn new(val: i32, expression: &str, status: i32) -> CellRef {
        Rc::new(RefCell::new(Self {
            val,
            expression: expression.chars().take(MAX_INPUT_LEN_CELL).collect(),
            status,
            dependencies: None,
            dependents: None,
        }))
    }
}
