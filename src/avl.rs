use std::cell::RefCell;
use std::cmp::max;
use std::rc::Rc;
use crate::cell::*;
// use crate::stack::*;
pub type Link = Option<Rc<RefCell<AvlNode>>>;
// type CellRef = Rc<RefCell<Cell>>;
// pub type Sheet = Vec<Vec<Rc<RefCell<Cell>>>>;

// #[derive(Clone)]
// pub struct Cell {
//     pub value: i32,
// }

pub struct SheetData {
    pub sheet: Vec<Vec<CellRef>>,
    pub flat: Vec<CellRef>,
}

impl SheetData {
    pub fn new(rows: usize, cols: usize) -> Self {
        let mut flat: Vec<CellRef> = Vec::with_capacity(rows * cols);
        for _ in 0..(rows * cols) {
            flat.push(Cell::new(0, "", 0));
        }

        let mut sheet: Vec<Vec<CellRef>> = Vec::with_capacity(rows);
        for i in 0..rows {
            let start = i * cols;
            let end = start + cols;
            sheet.push(flat[start..end].to_vec());
        }

        SheetData { sheet, flat }
    }

    pub fn get(&self, row: usize, col: usize) -> CellRef {
        self.sheet[row][col].clone()
    }

    pub fn calculate_row_col(&self, target: &CellRef) -> Option<(usize, usize)> {
        self.flat.iter().position(|c| Rc::ptr_eq(c, target))
            .map(|i| (i / self.sheet[0].len(), i % self.sheet[0].len()))
    }
}

pub struct AvlNode {
    pub cell: Rc<RefCell<Cell>>,
    pub left: Link,
    pub right: Link,
    pub height: i32,
}

impl AvlNode {
    pub fn new(cell: Rc<RefCell<Cell>>) -> Rc<RefCell<Self>> {
        Rc::new(RefCell::new(Self {
            cell,
            left: None,
            right: None,
            height: 1,
        }))
    }
}

// fn calculate_row_col(cell: &Rc<RefCell<Cell>>, sheet: &Sheet) -> Option<(usize, usize)> {
//     for (i, row) in sheet.iter().enumerate() {
//         for (j, c) in row.iter().enumerate() {
//             if Rc::ptr_eq(cell, c) {
//                 return Some((i, j));
//             }
//         }
//     }
//     None
// }

fn compare_cells(a: &Rc<RefCell<Cell>>, b: &Rc<RefCell<Cell>>, sheet_data: &SheetData) -> std::cmp::Ordering {
    let (a_row, a_col) = sheet_data.calculate_row_col(a).unwrap();
    let (b_row, b_col) = sheet_data.calculate_row_col(b).unwrap();
    // let (b_row, b_col) = calculate_row_col(b, sheet).unwrap();

    match a_row.cmp(&b_row) {
        std::cmp::Ordering::Equal => a_col.cmp(&b_col),
        ord => ord,
    }
}

fn height(node: &Link) -> i32 {
    node.as_ref().map_or(0, |n| n.borrow().height)
}

fn get_balance(node: &Rc<RefCell<AvlNode>>) -> i32 {
    height(&node.borrow().left) - height(&node.borrow().right)
}

fn rotate_right(y: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let x = {
        let mut y_borrow = y.borrow_mut();
        y_borrow.left.take().unwrap()
    };
    let t2 = {
        let mut x_borrow = x.borrow_mut();
        x_borrow.right.take()
    };

    {
        let mut y_borrow = y.borrow_mut();
        y_borrow.left = t2;
    }
    {
        let mut x_borrow = x.borrow_mut();
        x_borrow.right = Some(y.clone());
    }
    {
        let mut y_borrow = y.borrow_mut();
        y_borrow.height = 1 + max(height(&y_borrow.left), height(&y_borrow.right));
    }
    {
        let mut x_borrow = x.borrow_mut();
        x_borrow.height = 1 + max(height(&x_borrow.left), height(&x_borrow.right));
    }
    x
}

fn rotate_left(x: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let y = {
        let mut x_borrow = x.borrow_mut();
        x_borrow.right.take().unwrap()
    };
    let t2 = {
        let mut y_borrow = y.borrow_mut();
        y_borrow.left.take()
    };

    {
        let mut x_borrow = x.borrow_mut();
        x_borrow.right = t2;
    }
    {
        let mut y_borrow = y.borrow_mut();
        y_borrow.left = Some(x.clone());
    }
    {
        let mut x_borrow = x.borrow_mut();
        x_borrow.height = 1 + max(height(&x_borrow.left), height(&x_borrow.right));
    }
    {
        let mut y_borrow = y.borrow_mut();
        y_borrow.height = 1 + max(height(&y_borrow.left), height(&y_borrow.right));
    }
    y
}


pub fn insert(node: Link, cell: Rc<RefCell<Cell>>, sheet_data: &SheetData) -> Link {
    if let Some(n) = node {
        let cmp;
        {
            let n_borrow = n.borrow();
            cmp = compare_cells(&cell, &n_borrow.cell, sheet_data);
        }
        {
            let mut n_borrow = n.borrow_mut();
            if cmp == std::cmp::Ordering::Less {
                n_borrow.left = insert(n_borrow.left.clone(), cell.clone(), sheet_data);
            } else if cmp == std::cmp::Ordering::Greater {
                n_borrow.right = insert(n_borrow.right.clone(), cell.clone(), sheet_data);
            } else {
                return Some(n.clone()); // Duplicate
            }

            n_borrow.height = 1 + max(height(&n_borrow.left), height(&n_borrow.right));
        }

        let balance = get_balance(&n);
        // Clone once and reuse for comparisons
        let left = n.borrow().left.clone();
        let right = n.borrow().right.clone();

        // LL Case
        if balance > 1 && compare_cells(&cell, &left.as_ref().unwrap().borrow().cell, sheet_data) == std::cmp::Ordering::Less {
            return Some(rotate_right(n));
        }
        // RR Case
        if balance < -1 && compare_cells(&cell, &right.as_ref().unwrap().borrow().cell, sheet_data) == std::cmp::Ordering::Greater {
            return Some(rotate_left(n));
        }
        // LR Case
        if balance > 1 && compare_cells(&cell, &left.as_ref().unwrap().borrow().cell, sheet_data) == std::cmp::Ordering::Greater {
            let left_rotated = rotate_left(left.unwrap());
            n.borrow_mut().left = Some(left_rotated);
            return Some(rotate_right(n));
        }
        // RL Case
        if balance < -1 && compare_cells(&cell, &right.as_ref().unwrap().borrow().cell, sheet_data) == std::cmp::Ordering::Less {
            let right_rotated = rotate_right(right.unwrap());
            n.borrow_mut().right = Some(right_rotated);
            return Some(rotate_left(n));
        }

        Some(n)
    } else {
        Some(AvlNode::new(cell))
    }
}

pub fn find(node: &Link, row: usize, col: usize, sheet_data: &SheetData) -> Link {
    if let Some(n) = node {
        let (n_row, n_col) = sheet_data.calculate_row_col(&n.borrow().cell).unwrap();
        if (row, col) == (n_row, n_col) {
            Some(n.clone())
        } else if (row, col) < (n_row, n_col) {
            find(&n.borrow().left, row, col, sheet_data)
        } else {
            find(&n.borrow().right, row, col, sheet_data)
        }
    } else {
        None
    }
}

fn min_value_node(node: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let mut current = node;
    while let Some(left) = {
        let current_borrow = current.borrow(); // This borrow ends at the end of the block
        current_borrow.left.clone()
    } {
        current = left;
    }
    current
}

pub fn delete_node(root: Link, row: usize, col: usize, sheet_data: &SheetData) -> Link {
    if let Some(node) = root {
        let mut node_borrow = node.borrow_mut();
        // let (n_row, n_col) = calculate_row_col(&node_borrow.cell, sheet).unwrap();
        let (n_row, n_col) = sheet_data.calculate_row_col(&node_borrow.cell).unwrap();
        if (row, col) < (n_row, n_col) {
            node_borrow.left = delete_node(node_borrow.left.clone(), row, col, sheet_data);
        } else if (row, col) > (n_row, n_col) {
            node_borrow.right = delete_node(node_borrow.right.clone(), row, col, sheet_data);
        } else {
            // Node found
            if node_borrow.left.is_none() || node_borrow.right.is_none() {
                return node_borrow.left.clone().or(node_borrow.right.clone());
            } else {
                let temp = min_value_node(node_borrow.right.clone().unwrap());
                node_borrow.cell = temp.borrow().cell.clone();
                // let (t_row, t_col) = calculate_row_col(&temp.borrow().cell, sheet).unwrap();
                let (t_row, t_col) = sheet_data.calculate_row_col(&temp.borrow().cell).unwrap();
                node_borrow.right = delete_node(node_borrow.right.clone(), t_row, t_col, sheet_data);
            }
        }

        node_borrow.height = 1 + max(height(&node_borrow.left), height(&node_borrow.right));
        drop(node_borrow);

        let balance = get_balance(&node);
        let left = node.borrow().left.clone();
        let right = node.borrow().right.clone();

        if balance > 1 && get_balance(&left.as_ref().unwrap()) >= 0 {
            return Some(rotate_right(node));
        }

        if balance > 1 && get_balance(&left.as_ref().unwrap()) < 0 {
            let left_rotated = rotate_left(left.unwrap());
            node.borrow_mut().left = Some(left_rotated);
            return Some(rotate_right(node));
        }

        if balance < -1 && get_balance(&right.as_ref().unwrap()) <= 0 {
            return Some(rotate_left(node));
        }

        if balance < -1 && get_balance(&right.as_ref().unwrap()) > 0 {
            let right_rotated = rotate_right(right.unwrap());
            node.borrow_mut().right = Some(right_rotated);
            return Some(rotate_left(node));
        }

        Some(node)
    } else {
        None
    }
}

// pub fn inorder_traversal(root: &Link, sheet_data: &SheetData) {
//     fn traverse(node: &Link, sheet_data: &SheetData) {
//         if let Some(n) = node {
//             let n_borrow = n.borrow();
//             traverse(&n_borrow.left, sheet_data);

//             let (r, c) = sheet_data.calculate_row_col(&n_borrow.cell).unwrap();
//             let cell = n_borrow.cell.borrow();
//             println!("({r}, {c}) = {}", cell.val);

//             traverse(&n_borrow.right, sheet_data);
//         }
//     }
//     traverse(root, sheet_data);
// }

// pub fn find(node: &Link, value: i32) -> Link {
//     match node {
//         Some(n) => {
//             let n_borrow = n.borrow();
//             if value == n_borrow.cell.borrow().val {
//                 Some(n.clone())
//             } else if value < n_borrow.cell.borrow().val {
//                 find(&n_borrow.left, value)
//             } else {
//                 find(&n_borrow.right, value)
//             }
//         }
//         None => None,
//     }
// }

// pub fn insert(node: Link, cell: Rc<RefCell<Cell>>) -> Link {
//     if let Some(n) = node {
//         let mut n_borrow = n.borrow_mut();

//         if cell.borrow().val < n_borrow.cell.borrow().val {
//             n_borrow.left = insert(n_borrow.left.clone(), cell.clone());
//         } else if cell.borrow().val > n_borrow.cell.borrow().val {
//             n_borrow.right = insert(n_borrow.right.clone(), cell.clone());
//         } else {
//             return Some(n.clone()); // Duplicate values not allowed
//         }

//         n_borrow.height = 1 + max(height(&n_borrow.left), height(&n_borrow.right));
//         drop(n_borrow);

//         let balance = get_balance(&n);
//         let node = n.clone();

//         if balance > 1 && cell.borrow().val < node.borrow().left.as_ref().unwrap().borrow().cell.borrow().val {
//             return Some(rotate_right(node));
//         }
//         if balance < -1 && cell.borrow().val > node.borrow().right.as_ref().unwrap().borrow().cell.borrow().val {
//             return Some(rotate_left(node));
//         }
//         if balance > 1 && cell.borrow().val > node.borrow().left.as_ref().unwrap().borrow().cell.borrow().val {
//             let left_rotated = rotate_left(node.borrow().left.clone().unwrap());
//             node.borrow_mut().left = Some(left_rotated);
//             return Some(rotate_right(node));
//         }
//         if balance < -1 && cell.borrow().val < node.borrow().right.as_ref().unwrap().borrow().cell.borrow().val {
//             let right_rotated = rotate_right(node.borrow().right.clone().unwrap());
//             node.borrow_mut().right = Some(right_rotated);
//             return Some(rotate_left(node));
//         }

//         Some(node)
//     } else {
//         Some(AvlNode::new(cell))
//     }
// }

// pub fn delete_node(root: Link, value: i32) -> Link {
//     if let Some(node) = root {
//         let mut node_borrow = node.borrow_mut();

//         if value < node_borrow.cell.borrow().val {
//             node_borrow.left = delete_node(node_borrow.left.clone(), value);
//         } else if value > node_borrow.cell.borrow().val {
//             node_borrow.right = delete_node(node_borrow.right.clone(), value);
//         } else {
//             // Node to be deleted found
//             if node_borrow.left.is_none() || node_borrow.right.is_none() {
//                 return node_borrow.left.clone().or(node_borrow.right.clone());
//             } else {
//                 let temp = min_value_node(node_borrow.right.clone().unwrap());
//                 let replacement_value = temp.borrow().cell.borrow().val;
//                 node_borrow.cell = Rc::new(RefCell::new(Cell { value: replacement_value }));
//                 node_borrow.right = delete_node(node_borrow.right.clone(), replacement_value);
//             }
//         }

//         node_borrow.height = 1 + max(height(&node_borrow.left), height(&node_borrow.right));
//         drop(node_borrow);

//         let balance = get_balance(&node);

//         if balance > 1 && get_balance(&node.borrow().left.as_ref().unwrap()) >= 0 {
//             return Some(rotate_right(node));
//         }

//         if balance > 1 && get_balance(&node.borrow().left.as_ref().unwrap()) < 0 {
//             let left_rotated = rotate_left(node.borrow().left.clone().unwrap());
//             node.borrow_mut().left = Some(left_rotated);
//             return Some(rotate_right(node));
//         }

//         if balance < -1 && get_balance(&node.borrow().right.as_ref().unwrap()) <= 0 {
//             return Some(rotate_left(node));
//         }

//         if balance < -1 && get_balance(&node.borrow().right.as_ref().unwrap()) > 0 {
//             let right_rotated = rotate_right(node.borrow().right.clone().unwrap());
//             node.borrow_mut().right = Some(right_rotated);
//             return Some(rotate_left(node));
//         }

//         Some(node)
//     } else {
//         None
//     }
// }

// pub struct AvlTree {
//     root: Link,
// }

// impl AvlTree {
//     pub fn new() -> Self {
//         AvlTree { root: None }
//     }

//     pub fn insert(&mut self, value: i32) {
//         let cell = Rc::new(RefCell::new(Cell { value }));
//         self.root = insert(self.root.take(), cell);
//     }

//     pub fn delete(&mut self, value: i32) {
//         self.root = delete_node(self.root.take(), value);
//     }

//     pub fn find(&self, value: i32) -> bool {
//         find(&self.root, value).is_some()
//     }

//     pub fn inorder(&self) {
//         fn traverse(node: &Link) {
//             if let Some(n) = node {
//                 traverse(&n.borrow().left);
//                 print!("{} ", n.borrow().cell.borrow().val);
//                 traverse(&n.borrow().right);
//             }
//         }
//         traverse(&self.root);
//         println!();
//     }
// }