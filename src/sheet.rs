//! # SpreadSheet Program in Rust
//!
//! This crate implements a spreadsheet program in Rust, designed to handle
//! basic spreadsheet functionalities such as cell creation, deletion, and
//! evaluation of expressions. The program supports a grid-based layout where
//! each cell can contain a value or a formula. The program also includes
//! features for managing cell dependencies, detecting circular references,
//! and performing operations like SUM, AVG, MAX, MIN, and STDEV on ranges of
//! cells. The program is designed to be efficient and user-friendly, with
//! a focus on performance and ease of use.
use crate::avl::*;
use crate::cell::*;
use crate::stack::*;
use crate::extended::*;
use regex::Regex;
use std::time::Instant;
use std::env;
use std::io::{self, Write};
use std::rc::Rc;
use std::cell::RefCell;
use std::time::SystemTime;

use std::thread;
use std::time::Duration;

/// A static mutable variable to control the spreadsheet's output state.
/// When set to 1, output is enabled; otherwise, it is disabled.
pub static mut FLAG: i32 = 1;
/// A static mutable variable to store the number of rows in the spreadsheet.
pub static mut R: usize = 0;
/// A static mutable variable to store the number of columns in the spreadsheet.
pub static mut C: usize = 0;
/// A static mutable variable to store the starting row for displaying the spreadsheet.
pub static mut START_ROW: usize = 0;
/// A static mutable variable to store the starting column for displaying the spreadsheet.
pub static mut START_COL: usize = 0;
/// A static mutable variable to store the maximum length of input strings.
pub const MAX_INPUT_LEN: usize = 1000;

use lazy_static::lazy_static;

lazy_static! {
    static ref FUNC_REGEX: Regex = Regex::new(r"^([A-Z]{1,9})\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)(.*)$").unwrap();
    static ref SLEEP_REGEX_NUM: Regex = Regex::new(r"^SLEEP\((-?\d+)([^\)]*)\)$").unwrap();
    static ref SLEEP_REGEX_CELL: Regex = Regex::new(r"^SLEEP\(([A-Z]+)(\d+)([^\)]*)\)$").unwrap();
    static ref CELL_REF_REGEX: Regex = Regex::new(r"^([A-Z]+)(\d+)([^\n]*)$").unwrap();
}

/// Adds a dependency relationship from cell `c` to cell `dep` using an AVL tree.
/// 
/// # Arguments
/// * `dep` - The cell that depends on `c`
/// * `c` - The dependency cell
/// * `sheet_data` - The spreadsheet data structure
pub fn add_dependency(c: &CellRef, dep: &CellRef, sheet_data: &mut SheetData) {
    let existing_deps = {
        let cell = c.borrow();
        cell.dependencies.clone()
    };

    let new_deps = insert(existing_deps, Rc::clone(dep), sheet_data);

    c.borrow_mut().dependencies = new_deps;
}



/// Removes all dependencies from cells that depend on the specified `cell1`.
///
/// This is typically used when a cell's formula is changed or cleared,
/// and its dependents must be updated to reflect the removal of this dependency.
///
/// # Arguments
///
/// * `cell1` - The cell whose references should be removed from its dependents.
/// * `row` - The row index of `cell1`, used to locate the reference in other cells.
/// * `col` - The column index of `cell1`.
/// * `sheet_data` - A mutable reference to the spreadsheet data structure.
///
/// # How It Works
///
/// - Iteratively pops each dependent of `cell1`.
/// - For each dependent cell, removes the reference to `cell1` from its `dependencies` AVL tree.
/// - Ensures safe mutable access by using `take()` to temporarily extract values,
///   and then restoring ownership.
/// - Continues this process until no dependents remain.
///
pub fn delete_dependencies( row: usize, col: usize, sheet_data: &mut SheetData) {
    let cell1 = &sheet_data.sheet[row][col];
    loop {
        let dependent_node = {
            let mut cell_borrow = cell1.borrow_mut();
            match cell_borrow.dependents.take() {
                Some(node) => node,
                None => break, // exit loop if no more dependents
            }
        };
        let dependent_ref = dependent_node.borrow();
        let mut dependent = dependent_ref.cell.borrow_mut();
        dependent.dependencies = delete_node(dependent.dependencies.take(), row, col, sheet_data);

        pop_dependent(&cell1); // now it's safe to mutably borrow again
    }
}
/// Performs a depth-first search (DFS) to detect if a dependency path exists from the
/// `current` cell to the `target` cell in the spreadsheet graph.
/// 
/// This function is primarily used to detect **circular dependencies** between cells,
/// which would otherwise cause infinite evaluation loops.
/// 
/// # Arguments
///
/// * `current` - A reference to the cell where the DFS starts.
/// * `target` - A reference to the destination cell we are checking reachability for.
/// * `visited` - A bit-vector encoded as `Vec<u64>` to track visited cells efficiently.
/// * `current_row` - The row index of the `current` cell.
/// * `current_col` - The column index of the `current` cell.
/// * `sheet_data` - A reference to the entire spreadsheet's data structure for context.
///
/// # Returns
///
/// Returns `true` if a path exists from `current` to `target`, meaning
/// the `target` cell is reachable through dependencies — indicating a circular dependency.
/// Otherwise, returns `false`.
///
/// # How It Works
///
/// - Uses a bitwise visited map to avoid revisiting cells, based on their row-column index.
/// - If the target is directly in the dependencies of the current cell, it short-circuits.
/// - Otherwise, it traverses the dependency AVL tree recursively (in a stack-based manner).
///
pub fn dfs(
    current: &CellRef,
    target: &CellRef,
    visited: &mut Vec<u64>,
    current_row: usize,
    current_col: usize,
    sheet_data: &SheetData,
) -> bool {
    // Calculate bit indices for the visited ARRAY
    let index = current_row * unsafe { C } + current_col;
    let bit_index = index % 64;
    let vec_index = index / 64;
    
    // Early return if already visited
    if visited[vec_index] & (1 << bit_index) != 0 {
        return false;
    }
    
    // Mark as visited using bit operations
    visited[vec_index] |= 1 << bit_index;
    
    // Direct check first
    if Rc::ptr_eq(current, target) {
        return true;
    }
    
    // Target coordinates only need to be calculated once
    let (target_row, target_col) = sheet_data.calculate_row_col(target).unwrap_or((0, 0));
    
    // Check if direct dependency exists (faster than traversal)
    let cur = current.borrow();
    if find(&cur.dependencies, target_row, target_col, sheet_data).is_some() {
        return true;
    }
    
    // Use non-recursive stack-based traversal for better performance
    let mut stack = vec![cur.dependencies.clone()];
    while let Some(Some(node)) = stack.pop() {
        let dep_cell = &node.borrow().cell;
        let (dep_row, dep_col) = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0));
        
        if Rc::ptr_eq(dep_cell, target) ||
            (dep_row == target_row && dep_col == target_col) {
            return true;
        }
        
        // Check if dep_cell has been visited
        let dep_index = dep_row * unsafe { C } + dep_col;
        let dep_bit_index = dep_index % 64;
        let dep_vec_index = dep_index / 64;
        
        if visited[dep_vec_index] & (1 << dep_bit_index) == 0 {
            // Mark as visited
            visited[dep_vec_index] |= 1 << dep_bit_index;
            if dfs(dep_cell, target, visited, dep_row, dep_col, sheet_data) {
                return true;
            }
        }
        
        stack.push(node.borrow().left.clone());
        stack.push(node.borrow().right.clone());
    }
    
    false
}
/// Checks for the existence of a circular dependency between two cells in the spreadsheet.
///
/// This function determines whether a dependency path exists from `start` to `target`,
/// indicating a **cyclic reference**, which must be avoided in spreadsheet computations.
///
/// # Arguments
///
/// * `start` - The starting cell to begin the search from.
/// * `target` - The cell we want to check for being indirectly referenced by `start`.
/// * `start_row` - The row index of the `start` cell.
/// * `start_col` - The column index of the `start` cell.
/// * `sheet_data` - A reference to the complete spreadsheet structure.
///
/// # Returns
///
/// Returns `true` if a dependency path exists from `start` to `target`,
/// i.e., adding a reference from `target` to `start` would create a cycle.
/// Returns `false` otherwise.
///
/// # How It Works
///
/// - Initializes a `visited` bit-vector to keep track of explored cells.
/// - Calls [`dfs`] internally to perform a depth-first traversal through dependencies.
/// - Uses the `R` and `C` global constants to calculate bit indices for visited tracking.
pub fn check_loop(
    start: &CellRef,
    target: &CellRef,
    start_row: usize,
    start_col: usize,
    sheet_data: &SheetData,
) -> bool {
    // Quick check for direct self-reference
    if Rc::ptr_eq(start, target) {
        return true;
    }
    
    // Pre-calculate target position once
    let (target_row, target_col) = sheet_data.calculate_row_col(target).unwrap_or((0, 0));
    
    // Check if target is directly in start's dependencies (fast path)
    if find(&start.borrow().dependencies, target_row, target_col, sheet_data).is_some() {
        return true;
    }
    
    // Full dependency check
    let mut visited = vec![0u64; (unsafe { R * C }+63)/64];
    dfs(start, target, &mut visited, start_row, start_col, sheet_data)
}
/// Performs a depth-first search to check if any dependency of the current cell
/// lies within a specified rectangular range of cells.
///
/// This is useful when trying to detect if a formula indirectly refers
/// to any cell within a certain range, such as during bulk updates or validations.
///
/// # Arguments
///
/// * `current` - The cell to start the DFS from.
/// * `visited` - A boolean vector marking which cells have already been visited.
/// * `row1`, `col1` - The top-left corner of the target range.
/// * `row2`, `col2` - The bottom-right corner of the target range.
/// * `current_row`, `current_col` - The row and column of the current cell.
/// * `sheet_data` - A reference to the spreadsheet structure for cell access.
///
/// # Returns
///
/// Returns `true` if a path from `current` reaches any cell in the specified range;
/// otherwise, returns `false`.
///
/// # How It Works
///
/// - Checks if the current cell lies within the specified rectangular region.
/// - If not, traverses the `dependencies` AVL tree recursively to check
///   all downstream references.
/// - Marks visited cells to avoid redundant traversals.
pub fn dfs_range(
    current: &CellRef,
    visited: &mut Vec<bool>,
    row1: usize,
    col1: usize,
    row2: usize,
    col2: usize,
    current_row: usize,
    current_col: usize,
    // sheet: &mut Vec<Vec<CellRef>>,
    sheet_data: &SheetData,
) -> bool {
    if current_row >= row1 && current_row <= row2 && current_col >= col1 && current_col <= col2 {
        return true;
    }
    if !visited[current_row * unsafe { C } + current_col] {
        visited[current_row * unsafe { C } + current_col] = true;
        let cur = current.borrow();
        let mut stack = vec![cur.dependencies.clone()];
        while let Some(Some(node)) = stack.pop() {
            let dep_cell = &node.borrow().cell;
            // let dep_ptr = dep_cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
            // let dep_row = dep_ptr / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
            // let dep_col = dep_ptr / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
            let (dep_row , dep_col) = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0));
            // let dep_col = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0)).1;
            if dfs_range(
                dep_cell, visited, row1, col1, row2, col2, dep_row, dep_col, sheet_data,
            ) {
                return true;
            }
            stack.push(node.borrow().left.clone());
            stack.push(node.borrow().right.clone());
        }
    }
    false
}
/// Checks if the dependency graph from the `start` cell touches any cell within a rectangular range.
///
/// Used to detect potential **range-based cycles** or updates triggered
/// by a formula referencing a block of cells.
///
/// # Arguments
///
/// * `start` - The cell where the dependency check begins.
/// * `row1`, `col1` - Top-left cell of the range.
/// * `row2`, `col2` - Bottom-right cell of the range.
/// * `start_row`, `start_col` - Coordinates of the `start` cell.
/// * `sheet_data` - Reference to the spreadsheet’s data model.
///
/// # Returns
///
/// Returns `true` if any cell reachable from `start` is within the given range.
/// Otherwise, returns `false`.
///
/// # How It Works
///
/// - Initializes a `visited` vector for tracking cell visits.
/// - Calls [`dfs_range`] to perform a bounded DFS check against the range.
pub fn check_loop_range(
    start: &CellRef,
    row1: usize,
    col1: usize,
    row2: usize,
    col2: usize,
    start_row: usize,
    start_col: usize,
    // sheet: &mut Vec<Vec<CellRef>>,
    sheet_data: &SheetData,
) -> bool {
    let mut visited = vec![false; unsafe { R * C }];
    dfs_range(
        start,
        &mut visited,
        row1,
        col1,
        row2,
        col2,
        start_row,
        start_col,
        sheet_data,
    )
}
/// A utility function to perform depth-first traversal for topological sorting.
///
/// This function marks the current cell as visited, traverses all of its
/// dependencies recursively, and finally pushes the cell onto the stack. It ensures
/// that all cells it depends on are added to the stack before itself.
///
/// # Arguments
///
/// * `cell` - The current cell to process.
/// * `visited` - A mutable boolean vector that tracks whether a cell has already been visited.
/// * `sheet_data` - A reference to the full spreadsheet data structure.
/// * `stack` - A mutable reference to the stack where sorted cells are pushed.
///
/// # How It Works
///
/// - Calculates the index of the current cell in the 2D spreadsheet.
/// - If not visited:
///     - Marks the cell as visited.
///     - Recursively traverses all dependencies.
///     - Pushes the current cell to the result stack after its dependencies.
pub fn topological_sort_util(
    cell: &CellRef,
    visited: &mut Vec<bool>,
    sheet_data: &SheetData,
    stack: &mut StackLink,
) {
    if let Some((row, col)) = sheet_data.calculate_row_col(cell) {
        let index = row * unsafe { C } + col;

        // Skip if already visited
        if visited[index] {
            return;
        }
        
        visited[index] = true;

        // Use iterative approach instead of recursion for better performance
        let mut dep_stack = vec![(cell.borrow().dependencies.clone(), false)];
        
        while let Some((node_link, processed)) = dep_stack.pop() {
            if processed {
                // Node was processed, add to result stack
                if let Some(cell_node) = dep_stack.last() {
                    if let Some(ref node_rc) = cell_node.0 {
                        push(stack, Rc::clone(&node_rc.borrow().cell));
                    }
                }
                continue;
            }
            
            if let Some(node_rc) = node_link.clone() {
                let node = node_rc.borrow();
                
                // Mark this node for later processing after its dependencies
                dep_stack.push((node_link, true));
                
                // Process dependencies (right to left for stack order)
                if let Some(right) = node.right.clone() {
                    dep_stack.push((Some(right), false));
                }
                
                // Process the cell itself
                topological_sort_util(&node.cell, visited, sheet_data, stack);
                
                // Process left subtree
                if let Some(left) = node.left.clone() {
                    dep_stack.push((Some(left), false));
                }
            }
        }
        
        // Add the current cell to the stack
        push(stack, Rc::clone(cell));
    }
}

/// Initiates topological sorting from a given cell in the spreadsheet.
///
/// This function creates a new `visited` vector and starts a topological DFS traversal
/// from the given cell. The result is accumulated in the provided stack, with cells
/// ordered such that each cell appears after all of its dependencies.
///
/// # Arguments
///
/// * `start_cell` - The starting point for the topological sort.
/// * `sheet_data` - A reference to the spreadsheet’s internal state.
/// * `stack` - A mutable stack to which the sorted cells will be pushed in order.
pub fn topological_sort_from_cell(
    start_cell: &CellRef,
    sheet_data: &SheetData,
    stack: &mut StackLink,
) {
    // println!("Topological sort from cell");
    let mut visited = vec![false; unsafe { R * C }];
    topological_sort_util(start_cell, &mut visited, sheet_data, stack);
}
/// Handles scrolling logic for the spreadsheet view based on user input.
///
/// Adjusts the global viewport start positions (`START_ROW`, `START_COL`) to simulate
/// scrolling behavior in a terminal interface. Scrolling is done in blocks of 10 rows or columns.
///
/// # Arguments
///
/// * `input` - A string slice representing the scroll direction:
///     - `"w"`: Scroll up
///     - `"s"`: Scroll down
///     - `"a"`: Scroll left
///     - `"d"`: Scroll right
///
/// # Behavior
///
/// - Updates the global variables `START_ROW` and `START_COL` based on the direction.
/// - Ensures values remain within the bounds of the spreadsheet defined by `R` and `C`.
/// - Uses `saturating_sub` to safely handle potential underflows when scrolling near edges.
///
/// # Safety
///
/// This function uses `unsafe` to mutate static mutable variables, so it should be used
/// with caution and under the assumption of single-threaded context.
pub fn scroll(input: &str) -> i32 {
    unsafe {
        match input {
            "w" if START_ROW >= 10 => START_ROW -= 10,
            "w" => START_ROW = 0,
            "s" if START_ROW + 20 <= R - 1 => START_ROW += 10,
            "s" => START_ROW = R.saturating_sub(10),
            "a" if START_COL >= 10 => START_COL -= 10,
            "a" => START_COL = 0,
            "d" if START_COL + 20 <= C - 1 => START_COL += 10,
            "d" => START_COL = C.saturating_sub(10),
            _ => {}
        }
    }
    0
}
/// Pauses the execution of the program for a specified number of seconds.
/// 
/// This function is useful for simulating delays or waiting for a certain period
/// before proceeding with the next operation. It uses the `thread::sleep` function

pub fn sleep_seconds(seconds: u64) {
    thread::sleep(Duration::from_secs(seconds));
}
/// Converts a spreadsheet-style cell label (e.g., "B2", "AA10") into a (row, column) index.
///
/// Supports labels with up to 3 letters (A-Z) and up to 3 digits (0-9).
/// Valid labels must consist of uppercase letters followed by digits, with no interleaving.
///
/// # Arguments
///
/// * `label` - A string slice representing the cell label.
///
/// # Returns
///
/// * `Some((row, col))` if the label is valid.
/// * `None` if the label is invalid or out of bounds.
///
/// # Examples
///
/// ```
/// assert_eq!(label_to_index("A1"), Some((0, 0)));
/// assert_eq!(label_to_index("AA10"), Some((9, 26)));
/// assert_eq!(label_to_index("ZZZ999"), Some((998, 18277)));
/// assert_eq!(label_to_index("1A"), None); // invalid format
/// ```
pub fn label_to_index(label: &str) -> Option<(usize, usize)> {
    if label.len() > 6 || !label.chars().next().unwrap_or(' ').is_ascii_uppercase() {
        return None;
    }

    let mut count_letters = 0;
    let mut count_digits = 0;
    let mut alphabet = [0; 3];
    let mut number = [-1; 3];
    let chars: Vec<char> = label.chars().collect();
    let mut i = chars.len() as isize - 1;

    while i >= 0 {
        let ch = chars[i as usize];
        if ch.is_ascii_uppercase() {
            count_letters += 1;
            if count_digits == 0 {
                return None;
            }
            if count_letters > 3 {
                return None;
            }
            alphabet[3 - count_letters] = ch as usize - 'A' as usize + 1;
        } else if ch.is_ascii_digit() {
            count_digits += 1;
            if count_letters > 0 {
                return None;
            }
            if count_digits > 3 {
                return None;
            }
            number[3 - count_digits] = ch.to_digit(10).unwrap() as i32;
        } else {
            return None;
        }
        i -= 1;
    }

    if (number[0] == -1 && number[1] == -1 && number[2] == 0)
        || (number[0] == -1 && number[1] == 0)
        || (number[0] == 0)
    {
        return None;
    }

    if number[0] == -1 {
        number[0] = 0;
    }
    if number[1] == -1 {
        number[1] = 0;
    }

    let col = alphabet[2] + alphabet[1] * 26 + alphabet[0] * 26 * 26 - 1;
    let row = number[2] + number[1] * 10 + number[0] * 100 - 1;

    Some((row as usize, col as usize))
}
/// Converts a spreadsheet-style column label (e.g., "A", "AB", "ZZ") to a 0-based column index.
///
/// The label must be composed of only uppercase ASCII letters.
///
/// # Arguments
///
/// * `label` - A string slice representing the column label.
///
/// # Returns
///
/// * `Some(index)` if the label is valid.
/// * `None` if the label contains non-uppercase characters.
///
/// # Examples
///
/// ```
/// assert_eq!(col_label_to_index("A"), Some(0));
/// assert_eq!(col_label_to_index("Z"), Some(25));
/// assert_eq!(col_label_to_index("AA"), Some(26));
/// assert_eq!(col_label_to_index("AB"), Some(27));
/// assert_eq!(col_label_to_index("aB"), None); // lowercase not allowed
/// ```
pub fn col_label_to_index(label: &str) -> Option<usize> {
    if label.is_empty() {
        return None;
    }
    let mut index = 0;
    for ch in label.chars() {
        if !ch.is_ascii_uppercase() {
            return None;
        }
        index = index * 26 + (ch as usize - 'A' as usize + 1);
    }
    Some(index - 1)
}
/// Converts a 0-based column index into a spreadsheet-style column label (e.g., 0 → "A", 27 → "AB").
///
/// This is the inverse of `col_label_to_index`.
///
/// # Arguments
///
/// * `index` - A 0-based column index.
///
/// # Returns
///
/// * A `String` representing the column label.
///
/// # Examples
///
/// ```
/// assert_eq!(col_index_to_label(0), "A");
/// assert_eq!(col_index_to_label(25), "Z");
/// assert_eq!(col_index_to_label(26), "AA");
/// assert_eq!(col_index_to_label(27), "AB");
/// ```
pub fn col_index_to_label(mut index: usize) -> String {
    let mut buffer = ['\0'; 4];
    let mut i = 2;

    loop {
        buffer[i] = (b'A' + (index % 26) as u8) as char;
        index = index / 26;
        if index == 0 {
            break;
        }
        index -= 1;
        i -= 1;
    }
    buffer[i..=2].iter().collect()
}
/// Prints a 10x10 portion of the spreadsheet to the console starting from the current viewport (`START_ROW`, `START_COL`).
///
/// This function displays column labels at the top and row indices at the start of each row.
/// It prints cell values unless a cell has an error status (`status == 1`), in which case it prints `"ERR"`.
///
/// # Arguments
///
/// * `sheet` - A reference to a 2D vector of `CellRef`, representing the spreadsheet grid.
///
/// # Behavior
///
/// - Displays up to 10 rows and 10 columns from the current starting point.
/// - If `START_ROW + 10` or `START_COL + 10` exceed sheet dimensions, printing stops at the boundary.
/// - Uses `col_index_to_label` to display column headers (e.g., A, B, ..., Z, AA, AB...).
/// - Values are tab-separated for readability.
///
/// # Example Output
///
/// ```text
///     A       B       C       D       E       F       G       H       I       J
/// 1   42      15      0       23      ERR     4       7       9       2       5
/// 2   11      ERR     3       1       8       6       13      17      21      34
/// ...
/// ```
pub fn print_sheet(sheet: &Vec<Vec<CellRef>>) {
    unsafe {
        print!("\t");
        for col in START_COL..START_COL + 10 {
            if col >= C {
                break;
            }
            let label = col_index_to_label(col);
            print!("{}\t", label);
        }
        println!("");

        for row in START_ROW..START_ROW + 10 {
            if row >= R {
                break;
            }
            print!("{}\t", row + 1);
            for col in START_COL..START_COL + 10 {
                if col >= C {
                    break;
                }
                let cell = sheet[row][col].borrow();
                if cell.status == 1 {
                    print!("ERR\t");
                } else {
                    print!("{}\t", cell.val);
                }
            }
            println!("");
        }
    }
}
/// Splits a given string into a column label and a row number, if the string follows the format of a spreadsheet cell (e.g., "A1", "AB12").
///
/// This function separates the alphabetic part (representing the column label) and the numeric part (representing the row number) from a given input string.
/// If the string doesn't follow the valid format (such as containing letters after numbers, or invalid characters), the function returns `None`.
///
/// # Arguments
///
/// * `s` - A string slice that represents the cell reference (e.g., "A1", "AB12").
///
/// # Returns
///
/// Returns an `Option` containing a tuple `(label, number)` where:
/// - `label` is the column label (letters),
/// - `number` is the row number (digits).
///
/// Returns `None` if the input string is not a valid cell reference.
///
/// # Examples
///
/// ```rust
/// assert_eq!(split_label_and_number("A1"), Some(("A".to_string(), "1".to_string())));
/// assert_eq!(split_label_and_number("AB12"), Some(("AB".to_string(), "12".to_string())));
/// assert_eq!(split_label_and_number("A1B"), None);
/// ```
fn split_label_and_number(s: &str) -> Option<(String, String)> {
    let mut label = String::new();
    let mut number = String::new();
    for c in s.chars() {
        if c.is_ascii_alphabetic() {
            if !number.is_empty() {
                return None; // invalid format like A1B
            }
            label.push(c);
        } else if c.is_ascii_digit() {
            number.push(c);
        } else {
            return None;
        }
    }
    if label.is_empty() || number.is_empty() {
        None
    } else {
        Some((label, number))
    }
}
/// Evaluates a spreadsheet cell expression and updates the result value.
///
/// # Arguments
///
/// * `expr` - The string expression to evaluate.
/// * `rows` - Total number of rows in the spreadsheet.
/// * `cols` - Total number of columns in the spreadsheet.
/// * `sheet_data` - Mutable reference to the spreadsheet data structure.
/// * `result` - Mutable reference where the computed result will be stored.
/// * `row` - Current row index of the cell being evaluated.
/// * `col` - Current column index of the cell being evaluated.
/// * `call_value` - Flag indicating whether to update dependencies (1) or just evaluate (0).
///
/// # Return Value
///
/// Returns an integer status code:
/// * `0`: Success
/// * `-1`: Invalid expression
/// * `-2`: Division by Zero error to set status to 1
/// * `-4`: Circular dependency detected
///
/// # Functionality
///
/// This function parses and evaluates various types of spreadsheet expressions:
///
/// 1. **Simple numbers**: Direct integer values.
/// 2. **Basic arithmetic expressions**: Supports `+`, `-`, `*`, and `/` operations between numbers and cell references.
/// 3. **Cell references**: References to other cells in the format `A1`, `B2`, etc.
/// 4. **Range functions**: Functions operating on cell ranges:
///    * `SUM(A1:B3)`: Sum of all values in the range.
///    * `AVG(A1:B3)`: Average of all values in the range.
///    * `MAX(A1:B3)`: Maximum value in the range.
///    * `MIN(A1:B3)`: Minimum value in the range.
///    * `STDEV(A1:B3)`: Standard deviation of values in the range.
/// 5. **Special functions**:
///    * `SLEEP(n)`: Pauses execution for n seconds.
///    * `SLEEP(A1)`: Pauses execution for the number of seconds specified in cell A1.
///
/// The function also manages cell dependencies, tracking which cells depend on others to properly handle updates and detect circular references.
///
/// # How It Works
///
/// - Parses the expression to identify numbers, operators, and cell references.
/// - Evaluates the expression recursively, handling binary operations.
/// - Updates dependencies in the spreadsheet data structure.
/// - Checks for circular references using a depth-first search.
/// - Handles special cases like SUM, AVG, MAX, MIN, STDEV functions.
/// - Updates the result value and the cell's status accordingly.
pub fn evaluate_expression(
    expr: &str,
    rows: usize,
    cols: usize,
    sheet_data: &mut SheetData,
    result: &mut i32,
    row: &usize,
    col: &usize,
    call_value: i32,
) -> i32 {
    let mut count_status = 0;
    let mut col1: usize = 0;
    let mut row1: i32 = -1;
    let mut col2: usize = 0;
    let mut row2: i32 = -1;
    let value1 ;
    let value2 ;

    let trimmed_expr = expr.trim();
    // println!("trimmed_expr: {}", trimmed_expr);

    // Try to parse: just an integer
    if let Ok(val) = trimmed_expr.parse::<i32>() {
        *result = val;
        if call_value == 1 {
            delete_dependencies( *row, *col, sheet_data);
        }
        return 0;
    }
    let to_cell = &(sheet_data.sheet)[*row][*col].clone();
    if let Some(caps) = SLEEP_REGEX_NUM.captures(expr.trim())
    {
        let result_value = caps.get(1).unwrap().as_str().parse::<i32>().unwrap_or(-1);
        let temp = caps
            .get(2)
            .map_or(String::new(), |m| m.as_str().to_string());
        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }
        *result = result_value;
        
        if result_value < 0 {
            
            return 0; // Invalid sleep time
        }

        // Call sleep function (assuming a placeholder here)
        sleep_seconds(result_value.try_into().unwrap_or(0));
        return 0;
    }
    if let Some(caps) = SLEEP_REGEX_CELL.captures(expr.trim())
    {
        let label1 = caps.get(1).unwrap().as_str();
        let row1_str = caps.get(2).unwrap().as_str().to_string();
        let temp = caps
            .get(3)
            .map_or(String::new(), |m| m.as_str().to_string());

        // Validate that there are no extra characters after the number
        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }

        // Check for '0' in the cell reference
        if label1.chars().nth(label1.len()) == Some('0') {
            return -1; // Invalid cell
        }
        if row1_str.starts_with('0') {
            return -1; // Invalid expression
        }
        row1 = row1_str.parse::<i32>().unwrap_or(-1);
        row1 -= 1;
        if row1 < 0 {
            return -1; // Invalid cell
        }

        if let Some(val) = col_label_to_index(&label1) {
            col1 = val as usize;
        }

        // Validate cell boundaries
        if col1 >= cols || row1 >= rows as i32 {
            return -1; // Out-of-bounds error
        }

        // Check for circular dependency
        if check_loop(
            &(*sheet_data.sheet)[*row][*col],
            &(*sheet_data.sheet)[row1 as usize][col1],
            *row,
            *col,
            &*sheet_data,
        ) {
            return -4; // Circular dependency detected
        }

        // Check for errors in the referenced cell
        let mut count_status = 0;
        if (*(sheet_data.sheet))[row1 as usize][col1].borrow().status == 1 {
            count_status += 1; // Increment count if the referenced cell has an error
        }

        let result_value = (*(sheet_data.sheet))[row1 as usize ][col1].borrow().val;

        let from_cell = &(sheet_data.sheet)[row1 as usize][col1].clone();
        if call_value == 1 {
            // Delete old dependencies and add new ones
            // let current = (sheet_data.sheet)[*row][*col].clone();
            delete_dependencies( *row, *col, sheet_data);

            add_dependency(
                from_cell,
                &(sheet_data.sheet)[*row][*col].clone(),
                sheet_data,
            );
            push_dependent(
                &(sheet_data.sheet)[*row][*col],
                &(sheet_data.sheet)[row1 as usize][col1],
            );
        }
        *result = result_value;
        if count_status > 0 {
            return -2;
        }
        if result_value < 0 {
            return 0; // Invalid sleep time
        }
        sleep_seconds(result_value.try_into().unwrap_or(0));

        // If any dependents have errors, return -2
        

        return 0;
    }
    if let Some(op_i) = "+-*/".chars().find_map(|op| {
        trimmed_expr.find(op).map(|i| (i, op))
    }) {
        let (op_index, operator) = op_i;
        let (expr1, expr2) = trimmed_expr.split_at(op_index);
        let expr2 = &expr2[1..]; // skip operator

        let expr1 = expr1.trim();
        let expr2 = expr2.trim();

        
        // Process expr1
        if let Some((label1, num1)) = split_label_and_number(expr1) {
            if num1.starts_with('0') {
                return -1;
            }

            if let Some(val) = col_label_to_index(&label1) {
                col1 = val;
                if let Ok(r) = num1.parse::<i32>() {
                    row1 = r - 1;
                    if col1 >= cols || row1 < 0 || row1 >= rows as i32 {
                        return -1;
                    }

                    // Get reference to the cell
                    let cell1_ref = &(sheet_data.sheet)[row1 as usize][col1 as usize];
                    
                    // Check for cycles
                    if check_loop(&(sheet_data.sheet)[*row][*col], cell1_ref, *row, *col, sheet_data) {
                        return -4;
                    }
                    let cell = cell1_ref.borrow();
                    if cell.status == 1 {
                        count_status += 1;
                    }
                    value1 = cell.val;
                } else {
                    return -1;
                }
            } else {
                return -1;
            }
        } else if let Ok(val) = expr1.parse::<i32>() {
            value1 = val;
        } else {
            return -1;
        }

        // Process expr2
        if let Some((label2, num2)) = split_label_and_number(expr2) {
            if num2.starts_with('0') {
                return -1;
            }

            if let Some(val) = col_label_to_index(&label2) {
                col2 = val;
                if let Ok(r) = num2.parse::<i32>() {
                    row2 = r - 1;
                    if col2 >= cols || row2 < 0 || row2 >= rows as i32 {
                        return -1;
                    }

                    // Get reference to the cell
                    let cell2_ref = &(sheet_data.sheet)[row2 as usize][col2 as usize];
                    
                    // Check for cycles
                    if check_loop(&(sheet_data.sheet)[*row][*col], cell2_ref, *row, *col, sheet_data) {
                        return -4;
                    }
                    let cell = cell2_ref.borrow();
                    if cell.status == 1 {
                        count_status += 1;
                    }
                    value2 = cell.val;
                } else {
                    return -1;
                }
            } else {
                return -1;
            }
        } else if let Ok(val) = expr2.parse::<i32>() {
            value2 = val;
        } else {
            return -1;
        }

        // Dependency logic
        if call_value == 1 {
            delete_dependencies( *row, *col, sheet_data);

            if row1 >= 0 {
                // let dep_cell1 = (sheet_data.sheet)[row1 as usize][col1 as usize].clone();
                let from_cell = &(sheet_data.sheet)[row1 as usize][col1 as usize].clone();
                add_dependency(from_cell,to_cell, sheet_data);
                push_dependent(&(sheet_data.sheet)[*row][*col], &(sheet_data.sheet)[row1 as usize][col1 as usize]);
            }

            if row2 >= 0 && (col2 != col1 || row2 != row1) {
                // let dep_cell2 = (sheet_data.sheet)[row2 as usize][col2 as usize].clone();
                let from_cell = &(sheet_data.sheet)[row2 as usize][col2 as usize].clone();
                add_dependency(from_cell,to_cell, sheet_data);
                push_dependent(&(sheet_data.sheet)[*row][*col], &(sheet_data.sheet)[row2 as usize][col2 as usize]);
            }
        }

        if count_status > 0 {
            return -2;
        }

        // Perform the calculation
        match operator {
            '+' => *result = value1 + value2,
            '-' => *result = value1 - value2,
            '*' => *result = value1 * value2,
            '/' => {
                if value2 == 0 {
                    return -2;
                }
                *result = value1 / value2;
            }
            _ => return -1,
        }
        return 0;
    }

    if let Some(caps) = FUNC_REGEX.captures(expr.trim()) {
        let func = caps.get(1).unwrap().as_str().to_string();
        let label1 = caps.get(2).unwrap().as_str().to_string();
        let row1_str = caps.get(3).unwrap().as_str().to_string();
        let label2 = caps.get(4).unwrap().as_str().to_string();
        let row2_str = caps.get(5).unwrap().as_str().to_string();
        let temp = caps.get(6).map_or(String::new(), |m| m.as_str().to_string());

        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }
        if (func != "SUM" && func != "AVG" && func != "MAX" && func != "MIN" && func != "STDEV")
            || (label1.len() > 3 || label2.len() > 3)
        {
            return -1; // Invalid function
        }

        if row1_str.starts_with('0') {
            return -1; // Invalid expression
        }
        row1 = row1_str.parse::<i32>().unwrap_or(-1);
        row2 = row2_str.parse::<i32>().unwrap_or(-1);
        if temp.is_empty() {
            // Check validity of row and label lengths
            let len_row1 = row1.to_string().len();
            let len_row2 = row2.to_string().len();

            if expr
                .chars()
                .nth(func.len() + label1.len() + 1 + len_row1 + 1 + label2.len())
                == Some('0')
            {
                return -1; // Invalid cell
            }
            if expr
                .chars()
                .nth(func.len() + label1.len() + 1 + len_row1 + 1 + label2.len() + len_row2)
                != Some(')')
            {
                return -1; // Invalid cell
            }

            if let Some(val) = col_label_to_index(&label1) {
                col1 = val as usize;
            }
            if let Some(val) = col_label_to_index(&label2) {
                col2 = val as usize;
            }
            row1 -= 1;
            row2 -= 1;

            if col1 >= cols
                || row1 < 0
                || row1 >= rows as i32
                || col2 >= cols
                || row2 < 0
                || row2 >= rows as i32
                || row2 < row1
                || col2 < col1
            {
                return -1; // Out-of-bounds error
            }

            if check_loop_range(
                &(sheet_data.sheet)[*row as usize][*col as usize],
                row1 as usize,
                col1,
                row2 as usize,
                col2,
                *row,
                *col,
                &*sheet_data,
            ) {
                return -4; // Circular dependency detected
            }

            // Handle SUM function
            if func == "SUM" {
                *result = 0;
                if call_value == 1 {
                    delete_dependencies(
                        *row,
                        *col,
                        sheet_data,
                    );
                }

                for i in row1..=row2 {
                    for j in col1..=col2 {
                        {
                            let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                            if cell.status == 1 {
                                count_status += 1;
                            }
                            *result += cell.val;
                        }
                        let from_cell = &(sheet_data.sheet)[i as usize][j as usize].clone();
                        if call_value == 1 {
                            add_dependency(
                                from_cell,
                                to_cell,
                                sheet_data,
                            );
                            push_dependent(
                                &(sheet_data.sheet)[*row as usize][*col as usize],
                                &(sheet_data.sheet)[i as usize][j as usize],
                            );
                        }
                    }
                }

                if count_status > 0 {
                    return -2; // Error in dependents
                }
                return 0;
            }

            // Handle AVG function
            if func == "AVG" {
                *result = 0;
                let mut count = 0;

                if call_value == 1 {
                    //let mut cell = sheet[*row as usize][*col as usize].borrow_mut();
                    delete_dependencies(
                        *row,
                        *col,
                        sheet_data,
                    );
                }

                for i in row1..=row2 {
                    for j in col1..=col2 {
                        {
                            let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                            if cell.status == 1 {
                                count_status += 1;
                            }
                            *result += cell.val;
                            count += 1;
                        }
                        let from_cell = &(sheet_data.sheet)[i as usize][j as usize].clone();
                        if call_value == 1 {
                            add_dependency(
                                from_cell,
                                to_cell,
                                sheet_data,
                            );
                            push_dependent(
                                &(sheet_data.sheet)[*row as usize][*col as usize],
                                &(sheet_data.sheet)[i as usize][j as usize],
                            );
                        }
                    }
                }

                *result /= count;

                if count_status > 0 {
                    return -2; // Error in dependents
                }
                return 0;
            }

            // Handle MAX function
            if func == "MAX" {
                // println!("Inside MAX");
                *result = i32::MIN;
                if call_value == 1 {
                    delete_dependencies(
                        *row,
                        *col,
                        sheet_data,
                    );
                }
                for i in row1..=row2 {
                    for j in col1..=col2 {
                        let from_cell = &(sheet_data.sheet)[i as usize][j as usize].clone();
                        if call_value == 1 {
                            add_dependency(
                                from_cell,
                                to_cell,
                                sheet_data,
                            );
                            push_dependent(
                                &(sheet_data.sheet)[*row as usize][*col as usize],
                                &(sheet_data.sheet)[i as usize][j as usize],
                            );
                        }

                        let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                        if cell.status == 1 {
                            count_status += 1;
                        }

                        *result = cell.val.max(*result);
                    }
                }

                if count_status > 0 {
                    return -2; // Error in dependents
                }
                return 0;
            }

            if func == "MIN" {
                *result = i32::MAX;
                if call_value == 1 {
                    delete_dependencies(
                        *row,
                        *col,
                        sheet_data,
                    );
                }

                for i in row1..=row2 {
                    for j in col1..=col2 {
                        {
                            let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                            if cell.status == 1 {
                                count_status += 1;
                            }
                            *result = cell.val.min(*result);
                        }
                        let from_cell = &(sheet_data.sheet)[i as usize][j as usize].clone();
                        if call_value == 1 {
                            add_dependency(
                                from_cell,
                                to_cell,
                                sheet_data,
                            );
                            push_dependent(
                                &(sheet_data.sheet)[*row as usize][*col as usize],
                                &(sheet_data.sheet)[i as usize][j as usize],
                            );
                        }
                    }
                }

                if count_status > 0 {
                    return -2; // Error in dependents
                }
                return 0;
            }

            // Handle STDEV function
            if func == "STDEV" {
                let mut sum = 0;
                let mut count = 0;
                if call_value == 1 {
                    delete_dependencies(
                        *row,
                        *col,
                        sheet_data,
                    );
                }

                for i in row1..=row2 {
                    for j in col1..=col2 {
                        {
                            let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                            if cell.status == 1 {
                                count_status += 1;
                            }
                            sum += cell.val;
                            count += 1;
                        }
                        let from_cell = &(sheet_data.sheet)[i as usize][j as usize].clone();
                        if call_value == 1 {
                            add_dependency(
                                from_cell,
                                to_cell,
                                sheet_data,
                            );
                            push_dependent(
                                &(sheet_data.sheet)[*row as usize][*col as usize],
                                &(sheet_data.sheet)[i as usize][j as usize],
                            );
                        }
                    }
                }

                let mean: i32 = sum / count;
                let mut variance: f64 = 0.0;

                for i in row1..=row2 {
                    for j in col1..=col2 {
                        variance += (((sheet_data.sheet)[i as usize][j as usize].borrow().val - mean).pow(2)) as f64;
                    }
                }

                variance /= count as f64;
                *result = variance.sqrt().round() as i32;

                if count_status > 0 {
                    return -2; // Error in dependents
                }
                return 0;
            }
        }
    }
    // println!("DEBUG3: {}", expr);

    
    if let Some(caps) = CELL_REF_REGEX.captures(expr.trim())
    {
        let label1 = caps.get(1).unwrap().as_str();
        let row1_str = caps.get(2).unwrap().as_str().to_string();
        let temp = caps
            .get(3)
            .map_or(String::new(), |m| m.as_str().to_string());

        // Check for invalid cell references if there's extra content
        if !temp.is_empty() {
            return -1; // Invalid cell
        }

        // Check for '0' in the cell reference
        if label1.chars().nth(label1.len()) == Some('0') {
            return -1; // Invalid cell
        }
        if row1_str.starts_with('0') {
            return -1; // Invalid expression
        }
        row1 = row1_str.parse::<i32>().unwrap_or(-1);
        row1 -= 1;
        if row1 < 0 {
            return -1; // Invalid cell
        }
        if let Some(val) = col_label_to_index(&label1) {
            col1 = val;
        }

        // Validate cell boundaries
        if col1 >= cols || row1 >= rows as i32 {
            return -1; // Out-of-bounds error
        }

        // Check for circular dependency
        if check_loop(
            &(*(sheet_data.sheet))[*row][*col],
            &(*(sheet_data.sheet))[row1 as usize][col1],
            *row,
            *col,
            &*sheet_data,
        ) {
            return -4; // Circular dependency detected
        }

        // Check if the referenced cell has an error (status = 1)
        let mut count_status = 0;
        // let cell = .borrow_mut();
        if (*(sheet_data.sheet))[row1 as usize][col1].borrow().status == 1 {
            count_status += 1; // Increment if the referenced cell has an error
        }

        *result = (*(sheet_data.sheet))[row1 as usize][col1].borrow().val;

        // Update dependencies if needed
        if call_value == 1 {
            // let current = (sheet_data.sheet)[*row][*col].clone();
            delete_dependencies( *row, *col, sheet_data);

            add_dependency(
                &(sheet_data.sheet)[row1 as usize][col1].clone(),
                &(sheet_data.sheet)[*row][*col].clone(),
                sheet_data,
            );
            push_dependent(
                &(sheet_data.sheet)[*row][*col],
                &(sheet_data.sheet)[row1 as usize][col1],
            );
        }

        // If any dependents have errors, return -2
        if count_status > 0 {
            return -2;
        }

        return 0; // Success
    }

    return -1;
}
/// Executes a command on the spreadsheet engine.
///
/// # Parameters
/// - `input`: A string representing the user command or cell assignment (e.g., `"A1=5"`, `"scroll_to B3"`, `"w"`, `"disable_output"`).
/// - `rows`: Total number of rows in the spreadsheet.
/// - `cols`: Total number of columns in the spreadsheet.
/// - `sheet_data`: Mutable reference to the spreadsheet data.
///
/// # Commands Supported
/// - `"q"`: Quit the program.
/// - `"w"`, `"s"`, `"a"`, `"d"`: Scroll the view.
/// - `"scroll_to <cell>"`: Scroll to a specific cell (e.g., `scroll_to B3`). Returns -1 if out of bounds or invalid format.
/// - `"disable_output"` / `"enable_output"`: Toggle output flag (controlled via unsafe global `FLAG`).
/// - `<cell>=<expression>`: Assign an expression to a cell (e.g., `A1=5`, `B2=A1+10`).
/// It performs the following:
///
/// 1. Splits the input into a label and an expression.
/// 2. Converts the label into a `(row, col)` index in the sheet.
/// 3. Validates the indices against the sheet size.
/// 4. Evaluates the expression.
/// 5. Updates the cell’s value and expression if successful.
/// 6. Triggers a topological sort to re-evaluate all dependent cells.
/// 7. If the expression is invalid (e.g., circular dependency), it marks the cell and
///    propagates the error status.
///
/// # Returns
/// - `0` on successful execution of most commands.
/// - `1` if the command is `"q"` (quit).
/// - `-1` for invalid commands, out-of-bounds access, or malformed input.
/// - `-2` if division by zero is attempted.
/// - `-4` if there is a circular dependency in expressions.

pub fn execute_command(input: &str, rows: usize, cols: usize, sheet_data: &mut SheetData) -> i32 {
    // Quick check for common commands
    match input {
        "q" => return 1,
        "w" | "s" | "a" | "d" => return scroll(input),
        "disable_output" => {
            unsafe { FLAG = 0; }
            return 0;
        },
        "enable_output" => {
            unsafe { FLAG = 1; }
            return 0;
        },
        _ => {}
    }
    // let mut col : usize = 0;
    // Optimize for scrolling command
    if input.starts_with("scroll_to ") {
        let captures = &input[10..]; // Skip "scroll_to " prefix
        let digit_pos = captures.find(|c: char| c.is_ascii_digit()).unwrap_or(0);
        let (col_label, row_str) = captures.split_at(digit_pos);
        
        if row_str.starts_with('0') {
            return -1;
        }
        
        let col = match col_label_to_index(col_label) {
            Some(val) => val,
            None => return -1,
        };
        
        let row = match row_str.parse::<usize>() {
            Ok(r) => r.saturating_sub(1),
            Err(_) => return -1,
        };
        
        if col >= cols || row >= rows {
            return -1;
        }
        
        unsafe {
            START_ROW = row;
            START_COL = col;
        }
        return 0;
    }
    
    // Cell assignment handling
    if let Some((label, expr)) = input.split_once('=') {
        let (row, col) = match label_to_index(label.trim()) {
            Some(rc) => rc,
            None => return -1,
        };
        
        if row >= rows || col >= cols {
            return -1;
        }
        
        let mut result = 0;
        let cell = (sheet_data.sheet)[row][col].clone();
        
        match evaluate_expression(expr.trim(), rows, cols, sheet_data, &mut result, &row, &col, 1) {
            0 | 1 => {
                // if sheet_data.sheet[row][col].borrow().occur == 0 {
                //     sheet_data.sheet[row][col].borrow_mut().occur += 1;
                // }
                // Update cell value and status
                {
                    let mut cell_mut = cell.borrow_mut();
                    cell_mut.val = result;
                    cell_mut.expression = expr.trim().to_string();
                    cell_mut.status = 0;
                }
                
                // Update dependents using topological sort
                let mut stack = None;
                topological_sort_from_cell(&cell, sheet_data, &mut stack);
                
                // Remove the current cell from stack since we just updated it
                pop(&mut stack);
                
                // Process dependents in topological order
                while let Some(dep_cell) = pop(&mut stack) {
                    if let Some((r, c)) = sheet_data.calculate_row_col(&dep_cell) {
                        // Avoid multiple borrows
                        let expr = dep_cell.borrow().expression.clone();
                
                        let mut res = 0;
                
                        match evaluate_expression(&expr, rows, cols, sheet_data, &mut res, &r, &c, 0) {
                            0 | 1 => {
                                let mut cell_mut = sheet_data.sheet[r][c].borrow_mut();
                                cell_mut.val = res;
                                cell_mut.status = 0;
                            },
                            -2 => {
                                sheet_data.sheet[r][c].borrow_mut().status = 1;
                            },
                            _ => {}
                        }
                    }
                }
                
                
                return 0;
            },
            -2 => {
                // if sheet_data.sheet[row][col].borrow().occur == 0 {
                //     sheet_data.sheet[row][col].borrow_mut().occur += 1;
                // }
                // Error in calculation
                // Update cell value and status
                {
                    let mut cell_mut = cell.borrow_mut();
                    cell_mut.expression = expr.trim().to_string();
                    cell_mut.status = 1;
                }
                
                // Update dependents using topological sort
                let mut stack = None;
                topological_sort_from_cell(&cell, sheet_data, &mut stack);
                
                // Skip current cell
                pop(&mut stack);
                
                // Process dependents
                while let Some(dep_cell) = pop(&mut stack) {
                    if let Some((r, c)) = sheet_data.calculate_row_col(&dep_cell) {
                        let expr = dep_cell.borrow().expression.clone();
                        let mut res = 0;
                        
                        match evaluate_expression(&expr, rows, cols, sheet_data, &mut res, &r, &c, 0) {
                            0 | 1 => {
                                let mut cell_mut = (sheet_data.sheet)[r][c].borrow_mut();
                                cell_mut.val = res;
                                cell_mut.status = 0;
                            },
                            -2 => (sheet_data.sheet)[r][c].borrow_mut().status = 1,
                            _ => {}
                        }
                    }
                }
                return -2;
            },
            code => return code, // Return error codes directly
        }
    }
    
    -1  // Invalid command
}

/// Entry point for the spreadsheet program.
///
/// This program initializes a spreadsheet with a specified number of rows and columns
/// passed as command-line arguments. It supports both a default mode and a special `-vim` mode
/// (triggering an extended version of the application).
///
/// # Command-Line Arguments
/// - `<rows>`: Number of rows in the spreadsheet (1 ≤ rows ≤ 999).
/// - `<columns>`: Number of columns in the spreadsheet (1 ≤ columns ≤ 18278).
/// - `-vim`: Optional flag to run in extended mode (`extended::run_extended()`).
///
/// # Behavior
/// - Parses arguments and validates input sizes.
/// - Initializes the spreadsheet data.
/// - Displays the spreadsheet initially and after each successful command (if output is enabled).
/// - Accepts commands in a loop via standard input.
/// - Processes commands using `execute_command`.
/// - Displays execution time and command result status (`ok`, `Loop Detected!`, or `Invalid Input`).
/// - Exits when `"q"` command is entered.

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() > 1 && args[1] == "-vim" {
        // Call the extended version's main function
        if let Err(err) = run_extended() {
            eprintln!("Error in extended mode: {}", err);
            std::process::exit(-1);
        }
        return;
    }

    if args.len() != 3 {
        eprintln!("Usage: {} <No. of rows> <No. of columns>", args[0]);
        std::process::exit(-1);
    }

    let r: usize = args[1].parse().unwrap_or_else(|_| {
        eprintln!("Invalid number for rows.");
        std::process::exit(-1);
    });

    let c: usize = args[2].parse().unwrap_or_else(|_| {
        eprintln!("Invalid number for columns.");
        std::process::exit(-1);
    });

    unsafe {
        R = r;
        C = c;
    }

    if r < 1 || r > 999 {
        eprintln!("Invalid Input < 1<=R<=999 >");
        std::process::exit(-1);
    }

    if c < 1 || c > 18278 {
        eprintln!("Invalid Input < 1<=C<=18278 >");
        std::process::exit(-1);
    }

    let start_time = SystemTime::now();
    let mut sheet_data = SheetData::new(r, c);
    // create_sheet(&mut sheet);
    print_sheet(&(sheet_data.sheet));

    let elapsed = start_time.elapsed().unwrap().as_secs_f64();
    print!("[{:.2}] (ok) > ", elapsed);
    io::stdout().flush().unwrap();

    let stdin = io::stdin();
    let mut input = String::with_capacity(MAX_INPUT_LEN);

    loop {
        input.clear();
        if stdin.read_line(&mut input).is_err() {
            break;
        }

        input = input.trim_end().to_string();
        let start = Instant::now();

        let status = unsafe { execute_command(&input, R, C, &mut sheet_data) };

        if status == 1 {
            break;
        }

        let time_taken = start.elapsed().as_secs_f64();

        unsafe {
            if FLAG == 1 {
                print_sheet(&(sheet_data.sheet));
            }
        }

        match status {
            0 | -2 => print!("[{:.8}] (ok) > ", time_taken),
            -4 => print!("[{:.2}] (Loop Detected!) > ", time_taken),
            _ => print!("[{:.2}] (Invalid Input) > ", time_taken),
        }

        io::stdout().flush().unwrap();
    }
}
