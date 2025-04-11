use std::cell::RefCell;
use std::rc::Rc;
use crate::cell::Cell;

/// Type alias for shared, mutable reference to a `Cell`
pub type CellRef = Rc<RefCell<Cell>>;

/// Type alias for a link in the stack
pub type StackLink = Option<Rc<RefCell<StackNode>>>;

/// Stack node structure
#[derive(Clone)]
pub struct StackNode {
    pub cell: CellRef,
    pub next: StackLink,
}

impl StackNode {
    pub fn new(cell: CellRef, next: StackLink) -> StackLink {
        Some(Rc::new(RefCell::new(StackNode { cell, next })))
    }
}

/// Push to the dependents stack of a cell
pub fn push_dependent(cell: &CellRef, dep: CellRef) {
    let mut c = cell.borrow_mut();
    let new_node = StackNode::new(dep, c.dependents.clone());
    c.dependents = new_node;
}

/// Pop from the dependents stack of a cell
pub fn pop_dependent(cell: &CellRef) -> Option<CellRef> {
    let mut c = cell.borrow_mut();
    let top = c.dependents.take()?;
    let top_ref = top.borrow();
    let next = top_ref.next.clone();
    let dep_cell = top_ref.cell.clone();
    c.dependents = next;
    Some(dep_cell)
}

/// Push to a stack
pub fn push(stack: &mut StackLink, cell: CellRef) {
    let new_node = StackNode::new(cell, stack.clone());
    *stack = new_node;
}

/// Pop from a stack
pub fn pop(stack: &mut StackLink) -> Option<CellRef> {
    let top = stack.take()?;
    let top_ref = top.borrow();
    let next = top_ref.next.clone();
    let cell = top_ref.cell.clone();
    *stack = next;
    Some(cell)
}

/// Clears the dependents stack of a cell
pub fn free_dependents(cell: &CellRef) {
    let mut c = cell.borrow_mut();
    c.dependents = None;
}
