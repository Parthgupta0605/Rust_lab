use std::cell::RefCell;
use std::rc::Rc;
use crate::cell::*;

/// Type alias for shared, mutable reference to a `Cell`
// pub type CellRef = Rc<RefCell<Cell>>;

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
// pub fn free_dependents(cell: &CellRef) {
//     let mut c = cell.borrow_mut();
//     c.dependents = None;
// }
/// Prints the contents of a StackLink
pub fn print_stack(stack: &StackLink, name: &str) {
    let mut current = stack.clone();
    let mut index = 0;

    println!("Contents of stack '{}':", name);

    while let Some(node) = current {
        let node_ref = node.borrow();
        let cell = node_ref.cell.borrow();
        println!(
            "  [{}] -> Cell(val: {}, expr: '{}')",
            index, cell.val, cell.expression
        );
        current = node_ref.next.clone();
        index += 1;
    }

    if index == 0 {
        println!("  (Stack is empty)");
    }
}

