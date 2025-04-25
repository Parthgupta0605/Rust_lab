//! # Stack Module 
//!
//! This module defines the `StackNode` struct and associated functions for managing a stack of cells.
//! /// This module provides functionality for managing dependencies between `Cell` objects
//! using a stack-based data structure. The stack is implemented as a linked list of
//! `StackNode` elements, where each node holds a reference to a `Cell` and a link
//! to the next node. It enables efficient tracking and manipulation of dependent cells
//! during evaluation or updates.
use std::cell::RefCell;
use std::rc::Rc;
use crate::cell::*;

/// Type alias for shared, mutable reference to a `Cell`
// pub type CellRef = Rc<RefCell<Cell>>;

/// Type alias for a link in the stack
/// 
/// This alias represents a stack node as an `Option<Rc<RefCell<StackNode>>>`. Each node in the stack
/// contains a `CellRef` (reference to a `Cell`) and a reference to the next node in the stack.
pub type StackLink = Option<Rc<RefCell<StackNode>>>;

/// Stack node structure
/// 
/// Represents a single node in a stack, where each node holds a reference to a `Cell` and points to
/// the next node in the stack.
#[derive(Clone)]
pub struct StackNode {
    pub cell: CellRef,
    pub next: StackLink,
}

impl StackNode {
     /// Creates a new `StackNode` with the provided `cell` and `next` node
    ///
    /// # Arguments
    /// * `cell` - A reference to the `Cell` to be stored in this node.
    /// * `next` - The next node in the stack.
    ///
    /// # Returns
    /// * A `StackLink` (a wrapped `StackNode`) containing the provided `cell` and `next` node.
    pub fn new(cell: CellRef, next: StackLink) -> StackLink {
        Some(Rc::new(RefCell::new(StackNode { cell, next })))
    }
}

/// Push a dependent cell onto the dependents stack of the given `cell`.
///
/// This function adds a `dep` (dependent cell) onto the stack of dependents for the `cell`. The dependent
/// will be the first one in the stack (LIFO order).
/// 
/// # Arguments
/// * `cell` - A reference to the `Cell` that will have a dependent pushed onto its stack.
/// * `dep` - A reference to the `Cell` that will be added to the dependents stack.
pub fn push_dependent(cell: &CellRef, dep: &CellRef) {
    let mut c = cell.borrow_mut();
    let new_node = StackNode::new(dep.clone(), c.dependents.clone());
    c.dependents = new_node;
}

/// Pop a dependent cell from the dependents stack of the given `cell`.
///
/// This function removes and returns the top cell from the `cell`'s dependents stack. The cell at the top
/// of the stack (LIFO order) is returned, and the stack is updated to reflect this change.
///
/// # Arguments
/// * `cell` - A reference to the `Cell` whose dependents stack will be popped.
///
/// # Returns
/// * `Some(CellRef)` - The `Cell` reference of the dependent that was popped, if there is one.
/// * `None` - If the stack is empty.
pub fn pop_dependent(cell: &CellRef) -> Option<CellRef> {
    // let mut c = cell.borrow_mut();
    // let top = c.dependents.take()?;
    // let top_ref = top.borrow();
    // let next = top_ref.next.clone();
    // let dep_cell = top_ref.cell.clone();
    // c.dependents = next;
    // Some(dep_cell)
    let mut c = cell.borrow_mut();
    let top = c.dependents.take()?;

    // Narrow scope to drop top_ref before re-using c
    let (next, dep_cell) = {
        let top_ref = top.borrow();
        let next = top_ref.next.clone();
        let dep_cell = top_ref.cell.clone();
        (next, dep_cell)
    };

    c.dependents = next;
    Some(dep_cell)
}

/// Push a cell onto a stack.
///
/// This function pushes the provided `cell` onto the stack, making it the new top of the stack.
///
/// # Arguments
/// * `stack` - A mutable reference to the stack (`StackLink`) where the `cell` will be pushed.
/// * `cell` - The `CellRef` (reference to the `Cell`) to be added to the stack.
pub fn push(stack: &mut StackLink, cell: CellRef) {
    let new_node = StackNode::new(cell, stack.clone());
    *stack = new_node;
}

/// Pop a cell from the stack.
///
/// This function removes and returns the top cell from the stack. The stack is updated accordingly to reflect
/// the removal of the top node.
///
/// # Arguments
/// * `stack` - A mutable reference to the stack (`StackLink`) to pop from.
///
/// # Returns
/// * `Some(CellRef)` - The `Cell` reference of the cell that was popped, if the stack is not empty.
/// * `None` - If the stack is empty.
pub fn pop(stack: &mut StackLink) -> Option<CellRef> {
    let top = stack.take()?;
    let top_ref = top.borrow();
    let next = top_ref.next.clone();
    let cell = top_ref.cell.clone();
    *stack = next;
    Some(cell)
}