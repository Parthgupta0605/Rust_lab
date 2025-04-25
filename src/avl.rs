//! # AVL Tree implementation for a spreadsheet application
//! This module implements an AVL tree to manage cells in a spreadsheet.
//! The AVL tree is a self-balancing binary search tree, which ensures that the heights of the two child subtrees of any node differ by at most one.
//! This property makes AVL trees more efficient for lookups, insertions, and deletions compared to unbalanced binary search trees.

use std::cell::RefCell;
use std::cmp::max;
use std::rc::Rc;
use crate::cell::*;
/// Type alias for a reference to an AVL node
/// The `Link` type is an `Option` that can either be `Some` containing a reference to an `AvlNode` or `None`.
pub type Link = Option<Rc<RefCell<AvlNode>>>;
/// Represents the entire spreadsheet data structure.
///
/// `SheetData` stores a 2D grid of cells (`sheet`) and a flat 1D vector of all cells (`flat`)
/// to simplify certain operations like calculating the (row, col) of a specific cell reference.
/// 
/// This struct is used primarily by the AVL (dependency tracking / evaluation) system and
/// avoids importing higher-level logic from the `sheet` module to prevent circular dependencies.
pub struct SheetData {
      /// 2D matrix representation of the spreadsheet (rows x columns).
    pub sheet: Vec<Vec<CellRef>>,
     /// Flattened 1D vector of all cells in row-major order.
    /// Used for efficient indexing and lookups by position.
    pub flat: Vec<CellRef>,
}

impl SheetData {
     /// Creates a new `SheetData` instance with the given number of `rows` and `cols`.
    ///
    /// All cells are initialized with default values using `Cell::new(0, "", 0)`.
    /// The flat buffer is filled first, and then split into 2D rows in the `sheet` field.
    ///
    /// # Arguments
    /// * `rows` - The number of rows in the sheet.
    /// * `cols` - The number of columns in the sheet.
    ///
    /// # Returns
    /// A new `SheetData` instance with pre-allocated and linked cells.
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

       /// Returns a reference to a cell at a specific `(row, col)` in the sheet.
    ///
    /// # Arguments
    /// * `row` - Zero-based row index.
    /// * `col` - Zero-based column index.
    ///
    /// # Returns
    /// A `CellRef` (i.e., `Rc<RefCell<Cell>>`) pointing to the requested cell.
    ///
    /// # Panics
    /// Panics if the given `row` or `col` is out of bounds.
    pub fn get(&self, row: usize, col: usize) -> CellRef {
        self.sheet[row][col].clone()
    }
      /// Calculates the (row, col) position of a cell reference within the sheet.
    ///
    /// Searches through the flat list to find the index of the cell using `Rc::ptr_eq`,
    /// and maps that index back into a 2D `(row, col)` tuple.
    ///
    /// # Arguments
    /// * `target` - A reference to the cell whose position you want to find.
    ///
    /// # Returns
    /// * `Some((row, col))` if the cell exists in the sheet.
    /// * `None` if the cell is not part of this sheet.
    ///
    /// # Example
    /// ```
    /// let data = SheetData::new(3, 3);
    /// let cell = data.get(1, 2);
    /// assert_eq!(data.calculate_row_col(&cell), Some((1, 2)));
    /// ```
    pub fn calculate_row_col(&self, target: &CellRef) -> Option<(usize, usize)> {
        self.flat.iter().position(|c| Rc::ptr_eq(c, target))
            .map(|i| (i / self.sheet[0].len(), i % self.sheet[0].len()))
    }
}

/// Represents a node in an AVL tree used to track dependencies between spreadsheet cells.
///
/// Each node contains a reference to a cell (`CellRef`), as well as pointers to its
/// left and right children and its height in the tree. The AVL tree maintains balance
/// properties to ensure efficient insertions, deletions, and lookups.
pub struct AvlNode {
     /// A reference-counted, mutable reference to the cell associated with this node.   
    pub cell: Rc<RefCell<Cell>>,
    /// The left child in the AVL tree.
    pub left: Link,
     /// The right child in the AVL tree.
    pub right: Link,
     /// The height of the node in the AVL tree.
    ///
    /// Used to maintain balance during insertions and deletions
    pub height: i32,
}

impl AvlNode {
     /// Creates a new `AvlNode` with the given `CellRef` and initializes it as a leaf node.
    ///
    /// The node has no children (`left` and `right` are `None`) and starts with height `1`.
    ///
    /// # Arguments
    /// * `cell` - A reference-counted pointer to the `Cell` this node represents.
    ///
    /// # Returns
    /// A `Rc<RefCell<AvlNode>>`, allowing shared ownership and interior mutability of the node.
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

/// Compares two `CellRef`s based on their positions (row and column) in the spreadsheet.
///
/// This function is used for ordering cells in the AVL tree. Cells are compared first
/// by their row number, and then by their column number if the rows are equal.
///
/// # Arguments
/// * `a` - A reference to the first `CellRef`.
/// * `b` - A reference to the second `CellRef`.
/// * `sheet_data` - A reference to the `SheetData`, which is used to resolve the
///   row and column indices of the cells.
///
/// # Returns
/// An [`Ordering`](std::cmp::Ordering):  
/// - `Ordering::Less` if `a` comes before `b`,  
/// - `Ordering::Greater` if `a` comes after `b`,  
/// - `Ordering::Equal` if they are at the same position.
fn compare_cells(a: &Rc<RefCell<Cell>>, b: &Rc<RefCell<Cell>>, sheet_data: &SheetData) -> std::cmp::Ordering {
    let (a_row, a_col) = sheet_data.calculate_row_col(a).unwrap();
    let (b_row, b_col) = sheet_data.calculate_row_col(b).unwrap();

    match a_row.cmp(&b_row) {
        std::cmp::Ordering::Equal => a_col.cmp(&b_col),
        ord => ord,
    }
}

/// Returns the height of an AVL node.
///
/// This helper function safely retrieves the height of an AVL node,
/// returning `0` if the node is `None`.
///
/// # Arguments
/// * `node` - An optional reference to an `AvlNode`.
///
/// # Returns
/// The height of the node, or 0 if the node is `None`.
fn height(node: &Link) -> i32 {
    node.as_ref().map_or(0, |n| n.borrow().height)
}
/// Calculates the balance factor of an AVL node.
///
/// The balance factor is the difference in heights between the node's
/// left and right subtrees. It is used to determine whether the tree
/// needs to be rebalanced.
///
/// # Arguments
/// * `node` - A reference to an `Rc<RefCell<AvlNode>>` representing the AVL node.
fn get_balance(node: &Rc<RefCell<AvlNode>>) -> i32 {
    height(&node.borrow().left) - height(&node.borrow().right)
}
/// Performs a right rotation on an AVL node.
///
/// This operation is used to perform a right rotation on the given node `y` to
/// restore the balance of an AVL tree. The right rotation is typically used
/// when the left subtree of a node becomes too heavy (i.e., the balance factor
/// of the node is greater than 1).
///
/// # Arguments
/// * `y` - The `Rc<RefCell<AvlNode>>` representing the node to be rotated right.
///
/// # Returns
/// A new `Rc<RefCell<AvlNode>>` that represents the new root of the subtree
/// after the rotation.
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
/// Performs a left rotation on an AVL node.
///
/// This operation is used to perform a left rotation on the given node `x` to
/// restore the balance of an AVL tree. The left rotation is typically used
/// when the right subtree of a node becomes too heavy (i.e., the balance factor
/// of the node is less than -1).
///
/// # Arguments
/// * `x` - The `Rc<RefCell<AvlNode>>` representing the node to be rotated left.
///
/// # Returns
/// A new `Rc<RefCell<AvlNode>>` that represents the new root of the subtree
/// after the rotation.
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

/// Inserts a new `cell` into the AVL tree.
///
/// This function inserts a `cell` into the AVL tree while maintaining the balance of the tree. 
/// It ensures that the tree remains balanced after the insertion by applying the appropriate
/// rotations (right, left, or double rotations) when necessary.
///
/// # Arguments
/// * `node` - The root node of the AVL subtree to which the new `cell` should be inserted. This is a 
///            `Link` (i.e., an `Option<Rc<RefCell<AvlNode>>>`).
/// * `cell` - The `Rc<RefCell<Cell>>` representing the new cell to be inserted.
/// * `sheet_data` - A reference to the `SheetData`, which provides context for comparing cells.
///
/// # Returns
/// * `Link` - The updated root node of the AVL subtree after insertion and rebalancing. If the
///   node already contains the same `cell`, it returns the existing node (avoiding duplicates).
///
/// # Description                           
/// The function performs the following steps:
/// 1. It compares the `cell` with the `node`'s `cell` using the `compare_cells` function.
/// 2. It recursively traverses the AVL tree and inserts the `cell` in the correct position based on
///    the comparison result (less than or greater than).
/// 3. After insertion, it checks if the AVL tree needs rebalancing. If so, it applies the necessary
///    rotations (single or double rotations) to restore the balance.
/// 4. It updates the `height` of the nodes in the path of the inserted `cell` to ensure accurate height
///    calculations for subsequent balance checks.
///
/// The function considers the following balance cases for rebalancing:
/// - **LL (Left-Left Case):** A single right rotation is performed.
/// - **RR (Right-Right Case):** A single left rotation is performed.
/// - **LR (Left-Right Case):** A left rotation is followed by a right rotation.
/// - **RL (Right-Left Case):** A right rotation is followed by a left rotation.
///
pub fn insert(node: Link, cell: Rc<RefCell<Cell>>, sheet_data: &SheetData) -> Link {
    if let Some(n) = node {
        let cmp;
        {
            let n_borrow = n.borrow_mut();
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

/// Finds a node in the AVL tree corresponding to the given `row` and `col`.
///
/// This function searches the AVL tree for a specific `cell` based on its `row` and `col` values.
/// It compares the target `row` and `col` with the nodes' values using the `calculate_row_col` method from
/// the `SheetData` structure. If a match is found, it returns a reference to the corresponding `AvlNode`
/// wrapped in a `Some`, otherwise, it returns `None`.
///
/// # Arguments
/// * `node` - The root node of the AVL subtree in which to search. This is a `Link` (i.e., an `Option<Rc<RefCell<AvlNode>>>`).
/// * `row` - The row index of the `cell` to search for.
/// * `col` - The column index of the `cell` to search for.
/// * `sheet_data` - A reference to the `SheetData` structure, which is used to calculate the row and column indices of each `cell`.
///
/// # Returns
/// * `Link` - The `Link` (i.e., `Option<Rc<RefCell<AvlNode>>>`) of the node that corresponds to the given `row` and `col`.
///   If no such node exists, `None` is returned.
/// # Description
/// The function recursively traverses the AVL tree to locate the node that matches the given `row` and `col`:
/// 1. It compares the `row` and `col` of the target node with the current node's `row` and `col`.
/// 2. If a match is found, it returns the current node.
/// 3. If the target `row` and `col` are smaller than the current node's `row` and `col`, it recursively searches
///    the left subtree.
/// 4. If the target `row` and `col` are larger, it recursively searches the right subtree.
///
/// If the node does not exist in the tree, `None` is returned.
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

/// Finds the node with the minimum value in the AVL subtree rooted at the given `node`.
///
/// This function traverses the leftmost path in the AVL subtree, returning the node with the smallest
/// value (i.e., the leftmost node). It is typically used during the node deletion process in an AVL tree,
/// where the minimum node in the right subtree replaces the deleted node.
///
/// # Arguments
/// * `node` - The root node of the AVL subtree to search within. This is an `Rc<RefCell<AvlNode>>`.
///
/// # Returns
/// * `Rc<RefCell<AvlNode>>` - A reference-counted, mutable, borrowable `AvlNode` that contains the smallest
///   value in the subtree. This node is the leftmost node in the AVL tree.
///
/// # Description
/// The function iteratively traverses the left child of each node in the AVL subtree until it reaches
/// a node with no left child, which is the node with the smallest value. It then returns this node.
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
/// Deletes a node with the given `row` and `col` from the AVL tree.
///
/// This function deletes the node with the specified `row` and `col` from the AVL tree. It performs the
/// necessary rotations to maintain the AVL tree's balance after the deletion. If the node has two children,
/// it is replaced by its in-order successor (the node with the smallest value in the right subtree). If the node
/// has one or no children, it is removed directly.
///
/// # Arguments
/// * `root` - The root of the AVL tree to delete the node from. This is an `Option<Rc<RefCell<AvlNode>>>` (i.e., a `Link`).
/// * `row` - The row index of the node to delete.
/// * `col` - The column index of the node to delete.
/// * `sheet_data` - A reference to the `SheetData` structure used for calculating row and column indices of nodes.
///
/// # Returns
/// * `Link` - The new root of the AVL subtree after deletion. This is either a reference-counted pointer to the root node
///   (if the tree remains non-empty) or `None` (if the tree becomes empty).
///
/// # Description
/// The function works as follows:
/// 1. It first searches for the node to delete by comparing the `row` and `col` with the current node.
/// 2. If the node is found, it deletes it using the standard AVL deletion procedure:
///    - If the node has only one child or no children, it is removed directly.
///    - If the node has two children, it is replaced by its in-order successor (the smallest node in its right subtree).
/// 3. After the node is deleted, the tree is rebalanced if necessary by performing rotations.
///
/// The function uses left and right rotations as necessary to restore the AVL tree's balance factor after deletion.
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