mod avl;
mod cell;
mod stack;

use std::rc::Rc;
use std::cell::RefCell;
use crate::avl::{inorder_traversal, AvlNode, insert, delete_node, find, Link, Sheet};
use crate::cell::{Cell, CellRef};

fn make_sheet(rows: usize, cols: usize) -> Sheet {
    (0..rows)
        .map(|r| {
            (0..cols)
                .map(|c| {
                    let val = (r * cols + c) as i32;
                    Cell::new(val, "", 0)
                })
                .collect()
        })
        .collect()
}

fn get_cell(sheet: &Sheet, row: usize, col: usize) -> CellRef {
    sheet[row][col].clone()
}

fn print_found(node: Option<Rc<RefCell<AvlNode>>>, row: usize, col: usize) {
    match node {
        Some(n) => {
            let node_ref = n.borrow();                // <-- Extend this borrow's lifetime
            let cell_ref = node_ref.cell.borrow();    // <-- Now borrow the Cell
            println!("Cell at ({}, {}) = {}", row, col, cell_ref.val);
        }
        None => println!("Cell at ({}, {}) not found.", row, col),
    }
}

fn main() {
    let sheet = make_sheet(3, 3);
    let mut root: Link = None;

    println!("🟢 Inserting cells into AVL...");
    let c1 = get_cell(&sheet, 1, 2);
    let c2 = get_cell(&sheet, 0, 0);
    let c3 = get_cell(&sheet, 2, 1);

    root = insert(root, c2.clone(), &sheet);
    root = insert(root, c1.clone(), &sheet);   
    root = insert(root, c3.clone(), &sheet);

    println!("\n🔍 Finding inserted cells:");
    print_found(find(&root, 1, 2, &sheet), 1, 2);
    print_found(find(&root, 0, 0, &sheet), 0, 0);
    print_found(find(&root, 2, 1, &sheet), 2, 1);

    println!("\n🧾 Inorder traversal of AVL tree:");
    inorder_traversal(&root, &sheet);

    println!("\n🔄 Inserting more cells for testing balancing...");
    let c4 = get_cell(&sheet, 0, 1);
    let c5 = get_cell(&sheet, 0, 2);
    let c6 = get_cell(&sheet, 2, 2);
    
    root = insert(root, c4.clone(), &sheet);
    root = insert(root, c5.clone(), &sheet);
    root = insert(root, c6.clone(), &sheet);
    
    println!("\n🧾 Inorder traversal after insertions:");
    inorder_traversal(&root, &sheet);

    println!("\n🗑️  Deleting (1, 2)...");
    root = delete_node(root, 1, 2, &sheet);
    println!("\n🧾 Inorder traversal after deletion:");
    inorder_traversal(&root, &sheet);
    print_found(find(&root, 1, 2, &sheet), 1, 2);

    println!("\n🗑️  Deleting root node (0, 2)...");
    root = delete_node(root, 0, 2, &sheet);
    println!("\n🧾 Inorder traversal after deletion:");
    inorder_traversal(&root, &sheet);
    print_found(find(&root, 0, 2, &sheet), 0, 2);
}
