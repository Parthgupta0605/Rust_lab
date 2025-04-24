mod cell;
mod avl;
mod stack;
mod extended;
use crate::cell::*;
use crate::stack::*;
use crate::avl::*;

// use sscanf::sscanf;
use regex::Regex;
use std::time::Instant;
// use std::cell::RefCell;
use std::rc::Rc;
use std::time::SystemTime;
use std::env;
use std::io::{self, Write};

pub static mut FLAG: i32 = 1;
pub static mut R: usize = 0;
pub static mut C: usize = 0;
pub static mut START_ROW: usize = 0;
pub static mut START_COL: usize = 0;
pub const MAX_INPUT_LEN: usize = 1000;

// type CellRef = Rc<RefCell<Cell>>;

// pub fn create_sheet(sheet: &mut Vec<Vec<CellRef>>) {
//     unsafe {
//         sheet.clear(); // clear any existing content
//         sheet.reserve(R);

//         for _ in 0..R {
//             let mut row: Vec<CellRef> = Vec::with_capacity(C);
//             for _ in 0..C {
//                 row.push(Cell::new(0, "", 0));
//             }
//             sheet.push(row);
//         }
//     }
// }
// pub fn create_sheet(sheet: &mut Vec<Vec<CellRef>>) {
//     sheet.clear(); // clear any existing content

//     let mut flat: Vec<CellRef> = Vec::with_capacity(R * C);

//     for _ in 0..(R * C) {
//         flat.push(Cell::new(0, "", 0));
//     }

//     sheet.reserve(R);

//     for i in 0..R {
//         let start = i * C;
//         let end = start + C;

//         // Clone Rc pointers into each row
//         let row = flat[start..end].to_vec();
//         sheet.push(row);
//     }
// }

// pub struct SheetData {
//     pub sheet: Vec<Vec<CellRef>>,
//     pub flat: Vec<CellRef>,
// }

// impl SheetData {
//     pub fn new(rows: usize, cols: usize) -> Self {
//         let mut flat: Vec<CellRef> = Vec::with_capacity(rows * cols);
//         for _ in 0..(rows * cols) {
//             flat.push(Cell::new(0, "", 0));
//         }

//         let mut sheet: Vec<Vec<CellRef>> = Vec::with_capacity(rows);
//         for i in 0..rows {
//             let start = i * cols;
//             let end = start + cols;
//             sheet.push(flat[start..end].to_vec());
//         }

//         SheetData { sheet, flat }
//     }

//     pub fn get(&self, row: usize, col: usize) -> CellRef {
//         self.sheet[row][col].clone()
//     }

//     pub fn calculate_row_col(&self, target: &CellRef) -> Option<(usize, usize)> {
//         self.flat.iter().position(|c| Rc::ptr_eq(c, target))
//             .map(|i| (i / self.sheet[0].len(), i % self.sheet[0].len()))
//     }
// }

    pub fn add_dependency(c: CellRef, dep: CellRef, sheet_data: &mut SheetData) {
        // unsafe {
        //     c.borrow_mut().dependencies = insert(c.borrow_mut().dependencies.take(), dep.clone(), sheet_data);
        // }
        let mut cell = c.borrow_mut();
        let old_deps = cell.dependencies.take(); // take() gives you the inner value and sets it to None
        let new_deps = insert(old_deps, dep.clone(), sheet_data);
        cell.dependencies = new_deps;
    }

pub fn add_dependent(c: CellRef, dep: CellRef) {
    push_dependent(&c, dep);
}

// pub fn delete_dependencies(cell1: CellRef, row: usize, col: usize, sheet_data: &mut SheetData) {
//     while let Some(dependent_node) = cell1.borrow_mut().dependents.take() {
//         let dependent_ref = dependent_node.borrow();
//         let mut dependent = dependent_ref.cell.borrow_mut();
//         dependent.dependencies = delete_node(dependent.dependencies.take(), row, col, sheet_data);
//         pop_dependent(&cell1);
//     }
// }
pub fn delete_dependencies(cell1: CellRef, row: usize, col: usize, sheet_data: &mut SheetData) {
    loop {
        let (r,c) = sheet_data.calculate_row_col(&cell1).unwrap_or((0, 0));
        println!("Deleting dependencies for cell at row: {}, col: {}", r, c);
        // Move this into a scoped block to release borrow before pop_dependent
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

pub fn dfs(
    current: &CellRef,
    target: &CellRef,
    visited: &mut Vec<bool>,
    current_row: usize,
    current_col: usize,
    sheet_data: &SheetData,
) -> bool {
    let cur = current.borrow();
    if Rc::ptr_eq(current, target) {
        return true;
    }
    // let target_row = (target.as_ptr() as usize - sheet[0][0].as_ptr() as usize) / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
    // let target_col = (target.as_ptr() as usize - sheet[0][0].as_ptr() as usize) / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
    let target_row = sheet_data.calculate_row_col(target).unwrap_or((0, 0)).0;
    let target_col = sheet_data.calculate_row_col(target).unwrap_or((0, 0)).1;

    if find(&cur.dependencies, target_row, target_col, sheet_data).is_some() {
        return true;
    }
    if !visited[current_row * unsafe { C } + current_col] {
        visited[current_row * unsafe { C } + current_col] = true;
        let mut stack = vec![cur.dependencies.clone()];
        while let Some(Some(node)) = stack.pop() {
            let dep_cell = &node.borrow().cell;
            // let dep_ptr = dep_cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
            // let dep_row = dep_ptr / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
            // let dep_col = dep_ptr / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
            let dep_row = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0)).0;
            let dep_col = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0)).1;
            if dfs(dep_cell, target, visited, dep_row, dep_col, sheet_data) {
                return true;
            }
            stack.push(node.borrow().left.clone());
            stack.push(node.borrow().right.clone());
        }
    }
    false
}

pub fn check_loop(
    start: &CellRef,
    target: &CellRef,
    start_row: usize,
    start_col: usize,
    // target_row: usize,
    // target_col: usize,
    // sheet: &mut Vec<Vec<CellRef>>,
    sheet_data: &SheetData,
) -> bool {
    let mut visited = vec![false; unsafe { R * C }];
    let result = dfs(start, target, &mut visited, start_row, start_col, sheet_data);
    result
}

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
    sheet_data: &SheetData
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
            let dep_row = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0)).0;
            let dep_col = sheet_data.calculate_row_col(dep_cell).unwrap_or((0, 0)).1;
            if dfs_range(dep_cell, visited, row1, col1, row2, col2, dep_row, dep_col, sheet_data) {
                return true;
            }
            stack.push(node.borrow().left.clone());
            stack.push(node.borrow().right.clone());
        }
    }
    false
}

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
    dfs_range(start, &mut visited, row1, col1, row2, col2, start_row, start_col, sheet_data)
}

pub fn topological_sort_util(
    cell: &CellRef,
    visited: &mut Vec<bool>,
    sheet_data: &SheetData,
    stack: &mut StackLink,
) {
    if let Some((row, col)) = sheet_data.calculate_row_col(cell) {
        let index = row * unsafe { C } + col;

        if !visited[index] {
            visited[index] = true;

            // Clone dependencies to traverse safely
            let dependencies = cell.borrow().dependencies.clone();

            topological_dfs_traverse(dependencies, visited, sheet_data, stack);

            // After processing dependencies, push the current cell
            push(stack, Rc::clone(cell));
        }
    }
}

/// Helper recursive function to traverse AVL dependencies
fn topological_dfs_traverse(
    node_link: avl::Link,
    visited: &mut Vec<bool>,
    sheet_data: &SheetData,
    stack: &mut StackLink,
) {
    if let Some(node_rc) = node_link {
        let node = node_rc.borrow();

        // Visit the dependency cell
        topological_sort_util(&node.cell, visited, sheet_data, stack);

        // Traverse left subtree
        topological_dfs_traverse(node.left.clone(), visited, sheet_data, stack);

        // Traverse right subtree
        topological_dfs_traverse(node.right.clone(), visited, sheet_data, stack);
    }
}

// pub fn topological_sort_util(cell: &CellRef, visited: &mut Vec<bool>, sheet_data: &SheetData, stack: &mut StackLink) {
//     // let ptr_offset = cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
//     // let row = ptr_offset / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
//     // let col = ptr_offset / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
//     println!("Topological sort util for cell");
//     let row = sheet_data.calculate_row_col(cell).unwrap_or((0,0)).0;
//     let col = sheet_data.calculate_row_col(cell).unwrap_or((0,0)).1;

//     if !visited[row * unsafe { C } + col] {
//         visited[row * unsafe { C } + col] = true;
//         let cell_borrow = cell.borrow();
//         let mut deps_stack = vec![cell_borrow.dependencies.clone()];
//         while let Some(Some(node)) = deps_stack.pop() {
//             // println!("Node: {:?}", node.borrow().cell.borrow().expression);
//             topological_sort_util(&node.borrow().cell, visited, sheet_data, stack);
//             if let Some(node_left) = node.borrow().left.clone() {
//                 topological_sort_util(&node_left.borrow().cell, visited, sheet_data, stack);
//             }
//             if let Some(node_right) = node.borrow().right.clone() {
//                 topological_sort_util(&node_right.borrow().cell, visited, sheet_data, stack);
//             }
//         }
//         push(stack, Rc::clone(cell));
//     }
// }

pub fn topological_sort_from_cell(start_cell: &CellRef, sheet_data: &SheetData, stack: &mut StackLink) {
    println!("Topological sort from cell");
    let mut visited = vec![false; unsafe { R * C }];
    topological_sort_util(start_cell, &mut visited, sheet_data, stack);
}

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

pub fn sleep_seconds(seconds: u64) {
    let start_time = SystemTime::now();
    while SystemTime::now().duration_since(start_time).unwrap().as_secs() < seconds {
        // busy wait
    }
}

pub fn label_to_index(label: &str) -> Option<(usize, usize)> {
    if label.len() > 6 {
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

pub fn col_label_to_index(label: &str) -> Option<usize> {
    let mut index = 0;
    for ch in label.chars() {
        if !ch.is_ascii_uppercase() {
            return None;
        }
        index = index * 26 + (ch as usize - 'A' as usize + 1);
    }
    Some(index - 1)
}

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



fn evaluate_expression(
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
    let mut row1: i32 = 0;
    let mut col2: usize = 0;
    let mut row2: i32 = 0;
    let  value1 ;
    let  value2 ;

    let trimmed_expr = expr.trim();
    // println!("trimmed_expr: {}", trimmed_expr);

    // Try to parse: just an integer
    if let Ok(val) = trimmed_expr.parse::<i32>() {
        *result = val;
        if call_value == 1 {
            println!("Inside call_value == 1");
            let current = &(sheet_data.sheet)[*row][*col];
            delete_dependencies(current.clone(), *row, *col, sheet_data);
        }
        return 0;
    }

    // Try to parse: general binary expression: <expr1><op><expr2>
    let mut op_index = None;
    let mut operator = '\0';
    // println!("DEBUG1: {}", expr);
    for (i, c) in trimmed_expr.char_indices() {
        if "+-*/".contains(c) {
            op_index = Some(i);
            operator = c;
            break;
        }
    }

    if let Some(op_i) = op_index {
        let (expr1, expr2) = trimmed_expr.split_at(op_i);
        let expr2 = &expr2[1..]; // skip operator

        let expr1 = expr1.trim();
        let expr2 = expr2.trim();

        // Try expr1 as cell (like A1) or as number
        if let Some((label1, num1)) = split_label_and_number(expr1) {
            if num1.starts_with('0') {
                return -1;
            }

            if let Some(val) = col_label_to_index(&label1) {
                col1 = val;
            }
            if let Ok(r) = num1.parse::<i32>() {
                row1 = r - 1;
                if  col1 >= cols  || row1 < 0 || row1 >= rows as i32 {
                    return -1;
                }

                if check_loop(
                    &(sheet_data.sheet)[*row][*col],
                    &(sheet_data.sheet)[row1 as usize][col1 as usize],
                    *row,
                    *col,
                    // row1 as usize,
                    // col1 as usize,
                    &*sheet_data,
                ) {
                    return -4;
                }
                let cell = &(sheet_data.sheet)[row1 as usize][col1 as usize].borrow();
                if cell.status == 1 {
                    count_status += 1;
                }

                value1 = cell.val;
            } else {
                return -1;
            }
        } else if let Ok(val) = expr1.parse::<i32>() {
            value1 = val;
        } else {
            return -1;
        }

        // Try expr2 as cell (like B2) or as number
        if let Some((label2, num2)) = split_label_and_number(expr2) {
            if num2.starts_with('0') {
                return -1;
            }

            if let Some(val) = col_label_to_index(&label2) {
                col2 = val;
            }
            if let Ok(r) = num2.parse::<i32>() {
                row2 = r - 1;
                if  col2 >= cols  || row2 < 0 || row2 >= rows as i32 {
                    return -1;
                }

                if check_loop(
                    &(sheet_data.sheet)[*row][*col],
                    &(sheet_data.sheet)[row2 as usize][col2 as usize],
                    *row,
                    *col,
                    // row2 as usize,
                    // col2 as usize,
                    &*sheet_data,
                ) {
                    return -4;
                }
                let cell = &(sheet_data.sheet)[row2 as usize][col2 as usize].borrow();
                if cell.status == 1 {
                    count_status += 1;
                }

                value2 = cell.val;
            } else {
                return -1;
            }
        } else if let Ok(val) = expr2.parse::<i32>() {
            value2 = val;
        } else {
            return -1;
        }
        println!("value2: {}", value2);
        println!("call_value: {}", call_value);

        // Dependency logic
        if call_value == 1 {
            let current = (sheet_data.sheet)[*row][*col].clone();
            delete_dependencies(current.clone(), *row, *col, sheet_data);

            if  row1 >= 0 {
                add_dependency((sheet_data.sheet)[row1 as usize][col1 as usize].clone(), current.clone(), sheet_data);
                add_dependent(current.clone(), (sheet_data.sheet)[row1 as usize][col1 as usize].clone());
            }

            if  row2 >= 0 && (col2 != col1 || row2 != row1) {
                add_dependency((sheet_data.sheet)[row2 as usize][col2 as usize].clone(), current.clone(), sheet_data);
                add_dependent(current, (sheet_data.sheet)[row2 as usize][col2 as usize].clone());
            }
        }

        if count_status > 0 {
            println!("count_status: {}", count_status);
            return -2;
        }

        match operator {
            '+' => {
                *result = value1 + value2;
                return 0;
            }
            '-' => {
                *result = value1 - value2;
                return 0;
            }
            '*' => {
                *result = value1 * value2;
                return 0;
            }
            '/' => {
                if value2 == 0 {
                    return -2;
                }
                *result = value1 / value2;
                return 0;
            }
            _ => {
                return -1; // Invalid operator
            }
        }
    } 

    // let mut func = String::new(); 
    // let mut label1 = String::new();
    // let mut label2 = String::new();
    // let mut temp = String::new();

    // let mut row1_str = String::new();
    // let mut row2_str = String::new();

    // println!("DEBUG2: {}", expr);

    let func_regex = Regex::new(r"^([A-Z]{1,9})\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)(.*)$").unwrap();
    if let Some(caps) = func_regex.captures(expr.trim()) {
        let func = caps.get(1).unwrap().as_str().to_string();
        let label1 = caps.get(2).unwrap().as_str().to_string();
        let row1_str = caps.get(3).unwrap().as_str().to_string();
        let label2 = caps.get(4).unwrap().as_str().to_string();
        let row2_str = caps.get(5).unwrap().as_str().to_string();
        let temp = caps.get(6).map_or(String::new(), |m| m.as_str().to_string());
        
        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }
        if (func != "SUM" && func != "AVG" && func != "MAX" && func != "MIN" && func != "STDEV") || (label1.len() > 3 || label2.len() > 3) {
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

        if expr.chars().nth(func.len() + label1.len() + 1 + len_row1 + 1 + label2.len()) == Some('0') {
            return -1; // Invalid cell
        }
        if expr.chars().nth(func.len() + label1.len() + 1 + len_row1 + 1 + label2.len() + len_row2) != Some(')') {
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

        if  col1 >= cols || row1 < 0 || row1 >= rows as i32 ||  col2 >= cols || row2 < 0 || row2 >= rows as i32 || row2 < row1 || col2 < col1 {
            return -1; // Out-of-bounds error
        }

        if check_loop_range(&(sheet_data.sheet)[*row as usize][*col as usize], row1 as usize, col1, row2 as usize, col2, *row, *col, &*sheet_data) {
            return -4; // Circular dependency detected
        }

        // Handle SUM function
        if func == "SUM" {
            *result = 0;
            if call_value == 1 {
                delete_dependencies((sheet_data.sheet)[*row as usize][*col as usize].clone(), *row, *col, sheet_data);
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
                    if call_value == 1 {
                        add_dependency((sheet_data.sheet)[i as usize][j as usize].clone(), (sheet_data.sheet)[*row as usize][*col as usize].clone(), sheet_data);
                        add_dependent((sheet_data.sheet)[*row as usize][*col as usize].clone(), (sheet_data.sheet)[i as usize][j as usize].clone());
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
                delete_dependencies((sheet_data.sheet)[*row as usize][*col as usize].clone(), *row, *col, sheet_data);
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
                    if call_value == 1 {
                        add_dependency((sheet_data.sheet)[i as usize][j as usize].clone(), (sheet_data.sheet)[*row as usize][*col as usize].clone(), sheet_data);
                        add_dependent((sheet_data.sheet)[*row as usize][*col as usize].clone(), (sheet_data.sheet)[i as usize][j as usize].clone());
                    }
                }
            }

            *result /= count ;

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
                delete_dependencies((sheet_data.sheet)[*row as usize][*col as usize].clone(), *row, *col, sheet_data);
            }
            // println!("{:?}", row1);
            // println!("{:?}", row2);
            // println!("{:?}", col1);
            // println!("{:?}", col2);
            for i in row1..=row2 {
                for j in col1..=col2 {
                    if call_value == 1 {
                        add_dependency((sheet_data.sheet)[i as usize][j as usize].clone(), (sheet_data.sheet)[*row as usize][*col as usize].clone(), sheet_data);
                        add_dependent((sheet_data.sheet)[*row as usize][*col as usize].clone(), (sheet_data.sheet)[i as usize][j as usize].clone());
                    }

                    let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                    // println!("{:?}", cell.status);
                    if cell.status == 1 {
                        println!("i: {:?}", i);
                        println!("j: {:?}", j);
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

        // Handle MIN function
        if func == "MIN" {
            *result = i32::MAX;
            if call_value == 1 {
                delete_dependencies((sheet_data.sheet)[*row as usize][*col as usize].clone(), *row, *col, sheet_data);
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

                    if call_value == 1 {
                        add_dependency((sheet_data.sheet)[i as usize][j as usize].clone(), (sheet_data.sheet)[*row as usize][*col as usize].clone(), sheet_data);
                        add_dependent((sheet_data.sheet)[*row as usize][*col as usize].clone(), (sheet_data.sheet)[i as usize][j as usize].clone());
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
                delete_dependencies((sheet_data.sheet)[*row as usize][*col as usize].clone(), *row, *col, sheet_data);
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
                    if call_value == 1 {
                        add_dependency((sheet_data.sheet)[i as usize][j as usize].clone(), (sheet_data.sheet)[*row as usize][*col as usize].clone(), sheet_data);
                        add_dependent((sheet_data.sheet)[*row as usize][*col as usize].clone(), (sheet_data.sheet)[i as usize][j as usize].clone());
                    }

                   
                }
            }

            let mean: i32 = sum / count;
            let mut variance: f64 = 0.0;

            for i in row1..=row2 {
                for j in col1..=col2 {
                    let cell = (sheet_data.sheet)[i as usize][j as usize].borrow();
                    variance += (cell.val - mean).pow(2);
                }
            }

            variance /= count ;
            *result = variance.sqrt().round() as i32;

            if count_status > 0 {
                return -2; // Error in dependents
            }
            return 0;
        }
    }
}
    // println!("DEBUG3: {}", expr);
    if let Some(caps) = regex::Regex::new(r"^SLEEP\((\d+)([^\)]*)\)$").unwrap().captures(expr.trim()) {
        let result_value = caps.get(1).unwrap().as_str().parse::<i32>().unwrap_or(-1);
        let temp = caps.get(2).map_or(String::new(), |m| m.as_str().to_string());
    
        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }
    
        if result_value < 0 {
            *result = result_value;
            return 0; // Invalid sleep time
        }
    
        // Call sleep function (assuming a placeholder here)
        sleep_seconds(result_value.try_into().unwrap_or(0)); 
        return 0;
    }
    // println!("DEBUG4: {}", expr);
    // SLEEP with cell reference parsing
    if let Some(caps) = regex::Regex::new(r"^SLEEP\(([A-Z]+)(\d+)([^\)]*)\)$").unwrap().captures(expr.trim()) {
        let label1 = caps.get(1).unwrap().as_str();
        let row1: usize = caps.get(2).unwrap().as_str().parse().unwrap_or(0);
        let temp = caps.get(3).map_or(String::new(), |m| m.as_str().to_string());

        // Validate that there are no extra characters after the number
        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }

        // Check for '0' in the cell reference
        if label1.chars().nth(label1.len()) == Some('0') {
            return -1; // Invalid cell
        }

        if let Some(val) = col_label_to_index(&label1) {
            col1 = val as usize;
        }
        let row1 = row1 - 1; // Convert 1-based index to 0-based

        // Validate cell boundaries
        if  col1 >= cols || row1 >= rows as usize {
            return -1; // Out-of-bounds error
        }

        // Check for circular dependency
        if check_loop(&(*sheet_data.sheet)[*row][*col], &(*sheet_data.sheet)[row1][col1], *row, *col, &*sheet_data) {
            return -4; // Circular dependency detected
        }

        // Check for errors in the referenced cell
        let mut count_status = 0;
        if (*(sheet_data.sheet))[row1][col1].borrow().status == 1 {
            count_status += 1; // Increment count if the referenced cell has an error
        }

        let result_value = (*(sheet_data.sheet))[row1][col1].borrow().val;

        if call_value == 1 {
            // Delete old dependencies and add new ones
            let current = (sheet_data.sheet)[*row][*col].clone();
            delete_dependencies(current.clone(), *row, *col, sheet_data);

            add_dependency((*(sheet_data.sheet))[row1][col1].clone(), (*(sheet_data.sheet))[*row][*col].clone(), sheet_data);
            add_dependent((*(sheet_data.sheet))[*row][*col].clone(), (*(sheet_data.sheet))[row1][col1].clone());
        }
        *result = result_value;
        // if result_value < 0 {
            
        //     return 0; // Invalid sleep time
        // }

        // Sleep for the time indicated by the referenced cell's value
        sleep_seconds(result_value.try_into().unwrap_or(0));

        // If any dependents have errors, return -2
        if count_status > 0 {
            return -2;
        }

        return 0;
    }
    // println!("DEBUG: {}", expr);
    // Handle cell reference parsing (e.g., "A1", "B2", etc.)
    if let Some(caps) = regex::Regex::new(r"^([A-Z]+)(\d+)([^\n]*)$").unwrap().captures(expr.trim()) {
        let label1 = caps.get(1).unwrap().as_str();
        let row1: usize = caps.get(2).unwrap().as_str().parse().unwrap_or(0);
        let temp = caps.get(3).map_or(String::new(), |m| m.as_str().to_string());

        // Check for invalid cell references if there's extra content
        if !temp.is_empty() {
            return -1; // Invalid cell
        }

        // Check for '0' in the cell reference
        if label1.chars().nth(label1.len()) == Some('0') {
            return -1; // Invalid cell
        }

        if let Some(val) = col_label_to_index(&label1) {
            col1 = val ;
        }
        let row1 = row1 - 1; // Convert 1-based index to 0-based

        // Validate cell boundaries
        if  col1 >= cols || row1 >= rows as usize {
            return -1; // Out-of-bounds error
        }

        // Check for circular dependency
        if check_loop(&(*(sheet_data.sheet))[*row][*col], &(*(sheet_data.sheet))[row1][col1], *row, *col, &*sheet_data) {
            return -4; // Circular dependency detected
        }

        // Check if the referenced cell has an error (status = 1)
        let mut count_status = 0;
        
        if (*(sheet_data.sheet))[row1][col1].borrow().status == 1 {
            count_status += 1; // Increment if the referenced cell has an error
        }

        *result = (*(sheet_data.sheet))[row1][col1].borrow().val;

        // Update dependencies if needed
        if call_value == 1 {
            let current = (sheet_data.sheet)[*row][*col].clone();
            delete_dependencies(current.clone(), *row, *col, sheet_data);

            add_dependency((*(sheet_data.sheet))[row1][col1].clone(), (*(sheet_data.sheet))[*row][*col].clone(), sheet_data);
            add_dependent((*(sheet_data.sheet))[*row][*col].clone(), (*(sheet_data.sheet))[row1][col1].clone());
        }

        // If any dependents have errors, return -2
        if count_status > 0 {
            return -2;
        }

        return 0; // Success
    }

    return -1;


}




pub fn execute_command(input: &str, rows: usize, cols: usize,sheet_data: &mut SheetData) -> i32 {
    match input {
        "q" => {
            // Drop all memory (handled automatically in Rust)
            return 1;
        },
        "w" | "s" | "a" | "d" => {
            return scroll(input);
        },
        _ => {}
    }
    let mut col: usize = 0;
    // let mut row: usize = 0;
    if let Some(captures) = input.strip_prefix("scroll_to ") {
        let (col_label, row_str) = captures.trim().split_at(captures.find(|c: char| c.is_ascii_digit()).unwrap_or(0));
        if row_str.starts_with("0") { return -1; }
        if let Some(val) = col_label_to_index(&col_label) {
            col = val as usize;
        }
        let row: usize = row_str.parse().unwrap_or(0_usize).saturating_sub(1);

        if col >= cols || row >= rows  {
            return -1;
        }

        unsafe {
            START_ROW = row;
            START_COL = col;
        }
        return 0;
    }

    match input {
        "disable_output" => { unsafe { FLAG = 0; } return 0; },
        "enable_output" => { unsafe { FLAG = 1; } return 0; },
        _ => {}
    }

    if let Some((label, expr)) = input.split_once('=') {
        // println!("label: {}, expr: {}", label, expr);        
        let (row, col) = match label_to_index(label.trim()) {
            Some(rc) => rc,
            None => return -1,
        };
        // println!("label ki row: {}, col: {}", row, col);
        if row >= rows || col >= cols { return -1; }

        let mut result = 0;
        match evaluate_expression(expr.trim(), rows, cols, sheet_data, &mut result, &row, &col, 1) {
            0 | 1 => {
                let cell = &(sheet_data.sheet)[row][col];
                cell.borrow_mut().val = result;
                cell.borrow_mut().expression = expr.trim().to_string();
                cell.borrow_mut().status = 0;

                let mut stack = None;
                topological_sort_from_cell(cell, sheet_data, &mut stack);
                
                pop(&mut stack);
                print_stack(&stack, "MyStack");

                while let Some(cell) = pop(&mut stack) {
                    println!("Inside stack_node loop");
                    // let cell = &stack_node.borrow().cell;
                    if let Some((r, c)) = sheet_data.calculate_row_col(&cell) {
                        println!("Row: {}, Col: {}", r, c);
                        let mut res = 0;
                        let code = {
                            let cell_ref = cell.borrow();
                            evaluate_expression(&cell_ref.expression, rows, cols, sheet_data, &mut res, &r, &c, 0)
                        }; 
                        match code {
                            0|1 => { 
                                let mut cell_new = (sheet_data.sheet)[r][c].borrow_mut();
                                println!("res: {}", res);
                                cell_new.val = res; 
                                cell_new.status = 0; 
                            },
                            -2 => (sheet_data.sheet)[r][c].borrow_mut().status = 1,
                            _ => {}
                        }
                    }
                }
                return 0;
            },
            -2 => {
                let cell = &(sheet_data.sheet)[row][col];
                cell.borrow_mut().expression = expr.trim().to_string();
                cell.borrow_mut().status = 1;

                let mut stack:stack::StackLink = None;
                topological_sort_from_cell(&(sheet_data.sheet)[row][col], sheet_data,&mut stack);
                pop(&mut stack);
                for stack_node in stack {
                    if let Some((r, c)) = sheet_data.calculate_row_col(&stack_node.borrow().cell) {
                    let mut res = 0;
                    match evaluate_expression(&stack_node.borrow().cell.borrow().expression, rows, cols, sheet_data, &mut res, &r, &c, 0) {
                        0 | 1 => { (sheet_data.sheet)[r][c].borrow_mut().val = res; (sheet_data.sheet)[r][c].borrow_mut().status = 0; },
                        -2 => (sheet_data.sheet)[r][c].borrow_mut().status = 1,
                        _ => {}
                    }
                }
            }
                return -2;
            },
            -4 => return -4,
            -1 => return -1,
            _ => return -1, // KABHI AAEGA NHI
        }
    }

    -1
}


fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() > 1 && args[1] == "-vim" {
        // Call the extended version's main function
        if let Err(err) = extended::run_extended() {
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
        // println!("Input: {}", input);
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