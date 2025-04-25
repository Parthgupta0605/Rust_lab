//! # Cell Module
//!
//! This module defines the `Cell` struct, which represents a single cell in a spreadsheet.

use std::cell::RefCell;
use std::rc::Rc;

/// Maximum length allowed for an input expression
pub const MAX_INPUT_LEN_CELL: usize = 35;
/// A reference-counted, mutable reference to a `Cell`.
///
/// Used throughout the sheet to share ownership and allow internal mutability.
pub type CellRef = Rc<RefCell<Cell>>;


use crate::avl::{Link as AvlLink}; // Assuming `avl.rs` defines AVL tree
use crate::stack::{StackLink};   // Assuming `stack.rs` defines Stack

/// Represents a single cell in a spreadsheet.
///
/// Each cell stores a numeric value, an optional expression that defines its value,
/// a status code, and maintains both its dependents (cells it depends on) and 
/// depdependencies (cells that depend on it). Dependencies are stored using an AVL tree
/// for efficient lookup, while dependents are stored as a stack for quick updates.
///
/// The expression is stored as a `String`, and only up to [`MAX_INPUT_LEN_CELL`] characters are kept.
#[derive(Clone)]
pub struct Cell {
    /// The evaluated numeric value of the cell.
    pub val: i32,            
    /// The expression assigned to the cell (e.g., `=A1+B2`).
    ///
    /// Stored as a `String`, and trimmed to [`MAX_INPUT_LEN_CELL`] characters during creation.               // Value of the cell
    pub expression: String,                 // Expression stored as a String
    /// Status flag for the cell:
    /// * `0` => OK
    /// * `1` => ERR (Division by zero)
    pub status: i32,                        // Status to determine if it has ERR
     /// AVL tree storing references to all cells that depends on this cell.
    ///
    /// Useful for fast dependency resolution and loop detection.
    pub dependencies: AvlLink,              // AVL tree of dependencies
    /// Stack storing references to all cells that this cell depends on.
    ///
    /// Useful for quick updates and recalculations.
    pub dependents: StackLink,              // Stack of dependents
}

impl Cell {

    /// Creates a new `CellRef` (a reference-counted cell) with the given value, expression, and status.
    ///
    /// The expression string will be truncated to at most [`MAX_INPUT_LEN_CELL`] characters.
    ///
    /// # Arguments
    /// * `val` - Initial evaluated value of the cell.
    /// * `expression` - The expression associated with this cell (e.g., `"A1+B1"`).
    /// * `status` - The initial status of the cell (0 = OK, 1 = ERR).
    ///
    /// # Returns
    /// A `CellRef`, i.e., `Rc<RefCell<Cell>>`, allowing shared mutable access.
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
