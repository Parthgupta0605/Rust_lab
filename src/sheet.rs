use crate::cell::*;
use crate::stack::*;
use crate::avl::*;

use sscanf::sscanf;
use regex::Regex;
use std::time::Instant;
use std::cell::RefCell;
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

type CellRef = Rc<RefCell<Cell>>;

pub fn create_sheet(sheet: &mut Vec<Vec<CellRef>>) {
    unsafe {
        sheet.clear(); // clear any existing content
        sheet.reserve(R);

        for _ in 0..R {
            let mut row: Vec<CellRef> = Vec::with_capacity(C);
            for _ in 0..C {
                row.push(Cell::new(0, "", 0));
            }
            sheet.push(row);
        }
    }
}


pub fn add_dependency(c: &CellRef, dep: &CellRef, sheet: &mut Vec<Vec<CellRef>>) {
    unsafe {
        c.borrow_mut().dependencies = insert(c.borrow_mut().dependencies.take(), dep, sheet, C);
    }
}

pub fn add_dependent(c: &CellRef, dep: CellRef) {
    push_dependent(c, dep);
}

pub fn delete_dependencies(cell1: &CellRef, row: usize, col: usize, sheet: &mut Vec<Vec<CellRef>>) {
    while let Some(dependent_node) = cell1.dependents.take() {
        let mut dependent = dependent_node.cell.borrow_mut();
        dependent.dependencies = delete_node(dependent.dependencies.take(), row, col, sheet, unsafe { C });
        pop_dependent(cell1);
    }
}

pub fn dfs(
    current: &CellRef,
    target: &CellRef,
    visited: &mut Vec<bool>,
    current_row: usize,
    current_col: usize,
    sheet: &mut Vec<Vec<CellRef>>,
) -> bool {
    let cur = current.borrow();
    if Rc::ptr_eq(current, target) {
        return true;
    }
    let target_row = (target.as_ptr() as usize - sheet[0][0].as_ptr() as usize) / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
    let target_col = (target.as_ptr() as usize - sheet[0][0].as_ptr() as usize) / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
    if find(cur.dependencies.as_ref(), target_row, target_col, sheet, unsafe { C }).is_some() {
        return true;
    }
    if !visited[current_row * unsafe { C } + current_col] {
        visited[current_row * unsafe { C } + current_col] = true;
        let mut stack = vec![cur.dependencies.clone()];
        while let Some(Some(node)) = stack.pop() {
            let dep_cell = &node.borrow().cell;
            let dep_ptr = dep_cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
            let dep_row = dep_ptr / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
            let dep_col = dep_ptr / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
            if dfs(dep_cell, target, visited, dep_row, dep_col, sheet) {
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
    sheet: &mut Vec<Vec<CellRef>>,
) -> bool {
    let mut visited = vec![false; unsafe { R * C }];
    let result = dfs(start, target, &mut visited, start_row, start_col, sheet);
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
    sheet: &mut Vec<Vec<CellRef>>,
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
            let dep_ptr = dep_cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
            let dep_row = dep_ptr / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
            let dep_col = dep_ptr / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
            if dfs_range(dep_cell, visited, row1, col1, row2, col2, dep_row, dep_col, sheet) {
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
    sheet: &mut Vec<Vec<CellRef>>,
) -> bool {
    let mut visited = vec![false; unsafe { R * C }];
    dfs_range(start, &mut visited, row1, col1, row2, col2, start_row, start_col, sheet)
}

pub fn topological_sort_util(cell: &CellRef, visited: &mut Vec<bool>, sheet: &Vec<Vec<CellRef>>, stack: &mut StackLink) {
    let ptr_offset = cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
    let row = ptr_offset / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
    let col = ptr_offset / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };

    if !visited[row * unsafe { C } + col] {
        visited[row * unsafe { C } + col] = true;
        let cell_borrow = cell.borrow();
        let mut deps_stack = vec![cell_borrow.dependencies.clone()];
        while let Some(Some(node)) = deps_stack.pop() {
            topological_sort_util(&node.borrow().cell, visited, sheet, stack);
            deps_stack.push(node.borrow().left.clone());
            deps_stack.push(node.borrow().right.clone());
        }
        push(stack, Rc::clone(cell));
    }
}

pub fn topological_sort_from_cell(start_cell: &CellRef, sheet: &Vec<Vec<CellRef>>, stack: &mut StackLink) {
    let mut visited = vec![false; unsafe { R * C }];
    topological_sort_util(start_cell, &mut visited, sheet, stack);
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
    sheet: &mut Vec<Vec<CellRef>>,
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
    let mut value1 = 0;
    let mut value2 = 0;

    let trimmed_expr = expr.trim();

    // Try to parse: just an integer
    if let Ok(val) = trimmed_expr.parse::<i32>() {
        *result = val;
        if call_value == 1 {
            let current = &mut sheet[*row][*col];
            delete_dependencies(current, *row, *col, sheet);
        }
        return 0;
    }

    // Try to parse: general binary expression: <expr1><op><expr2>
    let mut op_index = None;
    let mut operator = '\0';

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
                if col1 < 0 || col1 >= cols  || row1 < 0 || row1 >= rows as i32 {
                    return -1;
                }

                if check_loop(
                    &sheet[*row][*col],
                    &sheet[row1 as usize][col1 as usize],
                    *row,
                    *col,
                    // row1 as usize,
                    // col1 as usize,
                    sheet,
                ) {
                    return -4;
                }
                let cell = &sheet[row1 as usize][col1 as usize].borrow();
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
                if col2 < 0 || col2 >= cols  || row2 < 0 || row2 >= rows as i32 {
                    return -1;
                }

                if check_loop(
                    &sheet[*row][*col],
                    &sheet[row2 as usize][col2 as usize],
                    *row,
                    *col,
                    // row2 as usize,
                    // col2 as usize,
                    sheet,
                ) {
                    return -4;
                }
                let cell = &sheet[row2 as usize][col2 as usize].borrow();
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

        // Dependency logic
        if call_value == 1 {
            let current = &sheet[*row][*col];
            delete_dependencies(current, *row, *col, sheet);

            if col1 >= 0 && row1 >= 0 {
                add_dependency(&sheet[row1 as usize][col1 as usize], current, sheet);
                add_dependent(current, sheet[row1 as usize][col1 as usize]);
            }

            if col2 >= 0 && row2 >= 0 && (col2 != col1 || row2 != row1) {
                add_dependency(&sheet[row2 as usize][col2 as usize], current, sheet);
                add_dependent(current, sheet[row2 as usize][col2 as usize]);
            }
        }

        if count_status > 0 {
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

    let mut func = String::new(); 
    let mut label1 = String::new();
    let mut label2 = String::new();
    let mut temp = String::new();

    if sscanf(expr, "%9[A-Z](%[A-Z]%d:%[A-Z]%d)%s", &mut func, &mut label1, &mut row1, &mut label2, &mut row2, &mut temp) == 5 {
        if expr.chars().nth(func.len() + label1.len() + 1).unwrap() == '0' {
            return -1; // Invalid expression
        }

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
        if let Some(val) = col_label_to_index(&label1) {
            col1 = val as usize;
        }
        row1 -= 1;
        row2 -= 1;

        if col1 < 0 || col1 >= cols || row1 < 0 || row1 >= rows as i32 || col2 < 0 || col2 >= cols || row2 < 0 || row2 >= rows as i32 || row2 < row1 || col2 < col1 {
            return -1; // Out-of-bounds error
        }

        if check_loop_range(sheet, &sheet[*row as usize][*col as usize], row1, col1, row2, col2) {
            return -4; // Circular dependency detected
        }

        // Handle SUM function
        if func == "SUM" {
            *result = 0;
            if call_value == 1 {
                delete_dependencies(&sheet[*row as usize][*col as usize], *row, *col, sheet);
            }

            for i in row1..=row2 {
                for j in col1..=col2 {
                    let cell = sheet[i as usize][j as usize].borrow();
                    if cell.status == 1 {
                        count_status += 1;
                    }
                    *result += cell.val;
                    if call_value == 1 {
                        add_dependency(&sheet[i as usize][j as usize], &sheet[*row as usize][*col as usize], sheet);
                        add_dependent(&sheet[*row as usize][*col as usize], sheet[i as usize][j as usize]);
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
                delete_dependencies(&sheet[*row as usize][*col as usize], *row, *col, sheet);
            }

            for i in row1..=row2 {
                for j in col1..=col2 {
                    let cell = sheet[i as usize][j as usize].borrow();
                    if cell.status == 1 {
                        count_status += 1;
                    }
                    *result += cell.val;
                    count += 1;
                    if call_value == 1 {
                        add_dependency(&sheet[i as usize][j as usize], &sheet[*row as usize][*col as usize], sheet);
                        add_dependent(&sheet[*row as usize][*col as usize], sheet[i as usize][j as usize]);
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
            *result = i32::MIN;
            if call_value == 1 {
                delete_dependencies(&sheet[*row as usize][*col as usize], *row, *col, sheet);
            }

            for i in row1..=row2 {
                for j in col1..=col2 {
                    if call_value == 1 {
                        add_dependency(&sheet[i as usize][j as usize], &sheet[*row as usize][*col as usize], sheet);
                        add_dependent(&sheet[*row as usize][*col as usize], sheet[i as usize][j as usize]);
                    }

                    let cell = sheet[i as usize][j as usize].borrow();
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

        // Handle MIN function
        if func == "MIN" {
            *result = i32::MAX;
            if call_value == 1 {
                delete_dependencies(&sheet[*row as usize][*col as usize], *row, *col, sheet);
            }

            for i in row1..=row2 {
                for j in col1..=col2 {
                    let cell = sheet[i as usize][j as usize].borrow();
                    if cell.status == 1 {
                        count_status += 1;
                    }

                    if call_value == 1 {
                        add_dependency(&sheet[i as usize][j as usize], &sheet[row][col], sheet);
                        add_dependent(&sheet[*row as usize][*col as usize], &sheet[i][j]);
                    }

                    *result = cell.val.min(*result);
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
                delete_dependencies(&sheet[*row as usize][*col as usize], *row, *col, sheet);
            }

            for i in row1..=row2 {
                for j in col1..=col2 {
                    let cell = sheet[i as usize][j as usize].borrow();
                    if cell.status == 1 {
                        count_status += 1;
                    }

                    if call_value == 1 {
                        add_dependency(&sheet[i as usize][j as usize], &sheet[row][col], sheet);
                        add_dependent(&sheet[*row as usize][*col as usize], &sheet[i][j]);
                    }

                    sum += cell.val;
                    count += 1;
                }
            }

            let mean = sum / count ;
            let mut variance: i32 = 0;

            for i in row1..=row2 {
                for j in col1..=col2 {
                    let cell = sheet[i as usize][j as usize].borrow();
                    variance += (cell.val - mean).powi(2);
                }
            }

            variance /= count ;
            *result = (variance as f64).sqrt() as i32;

            if count_status > 0 {
                return -2; // Error in dependents
            }
            return 0;
        }
    }

    if let Some(caps) = regex::Regex::new(r"^SLEEP\((\d+)([^\)]*)\)$").unwrap().captures(expr.trim()) {
        let result_value = caps.get(1).unwrap().as_str().parse::<i32>().unwrap_or(-1);
        let temp = caps.get(2).map_or(String::new(), |m| m.as_str().to_string());
    
        if !temp.is_empty() {
            return -1; // Invalid format if there's extra content after the number
        }
    
        if result_value < 0 {
            return -1; // Invalid sleep time
        }
    
        // Call sleep function (assuming a placeholder here)
        sleep_seconds(result_value); 
        return 0;
    }

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
        if col1 < 0 || col1 >= cols || row1 < 0 || row1 >= rows as usize {
            return -1; // Out-of-bounds error
        }

        // Check for circular dependency
        if check_loop(&(*sheet)[*row][*col], &(*sheet)[row1][col1], *row, *col, sheet) {
            return -4; // Circular dependency detected
        }

        // Check for errors in the referenced cell
        let mut count_status = 0;
        if (*sheet)[row1][col1].borrow().status == 1 {
            count_status += 1; // Increment count if the referenced cell has an error
        }

        let result_value = (*sheet)[row1][col1].borrow().val;

        if call_value == 1 {
            // Delete old dependencies and add new ones
            let current = &(*sheet)[*row][*col];
            delete_dependencies(current, *row, *col, sheet);
            add_dependency(&(*sheet)[row1][col1], &(*sheet)[*row][*col], sheet);
            add_dependent(&(*sheet)[*row][*col], (*sheet)[row1][col1]);
        }

        // Sleep for the time indicated by the referenced cell's value
        sleep_seconds(result_value);

        // If any dependents have errors, return -2
        if count_status > 0 {
            return -2;
        }

        return 0;
    }

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
        if col1 < 0 || col1 >= cols || row1 < 0 || row1 >= rows as usize {
            return -1; // Out-of-bounds error
        }

        // Check for circular dependency
        if check_loop(&(*sheet)[*row][*col], &(*sheet)[row1][col1], *row, *col, sheet) {
            return -4; // Circular dependency detected
        }

        // Check if the referenced cell has an error (status = 1)
        let mut count_status = 0;
        
        if (*sheet)[row1][col1].borrow().status == 1 {
            count_status += 1; // Increment if the referenced cell has an error
        }

        *result = (*sheet)[row1][col1].borrow().val;

        // Update dependencies if needed
        if call_value == 1 {
            let current = &(*sheet)[*row][*col];
            delete_dependencies(current, *row, *col, sheet);
            add_dependency(&(*sheet)[row1][col1], &(*sheet)[*row][*col], sheet);
            add_dependent(&(*sheet)[*row][*col], (*sheet)[row1][col1]);
        }

        // If any dependents have errors, return -2
        if count_status > 0 {
            return -2;
        }

        return 0; // Success
    }

    return -1;


}




pub fn execute_command(input: &str, rows: usize, cols: usize, sheet: &mut Vec<Vec<CellRef>>) -> i32 {
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
    let col: usize = 0;
    let row: usize = 0;
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
        let (row, col) = match label_to_index(label.trim()) {
            Some(rc) => rc,
            None => return -1,
        };
        if row >= rows || col >= cols { return -1; }

        let mut result = 0;
        match evaluate_expression(expr.trim(), rows, cols, sheet, &mut result, &row, &col, 1) {
            0 | 1 => {
                let cell = &sheet[row][col];
                cell.borrow().val = result;
                cell.borrow().expression = expr.trim().to_string();
                cell.borrow().status = 0;

                let mut stack = None;
                topological_sort_from_cell(cell, sheet, &mut stack);
                pop(stack);

                for cell in stack {
                    let r = cell.row;
                    let c = cell.col;
                    let mut res = 0;
                    match evaluate_expression(&cell.expression, rows, cols, sheet, &mut res, r, c, 0) {
                        1 => { sheet[r][c].borrow().val = res; sheet[r][c].borrow().status = 0; },
                        -2 => sheet[r][c].borrow().status = 1,
                        _ => {}
                    }
                }
                return 0;
            },
            -2 => {
                sheet[row][col].expression = expr.trim().to_string();
                sheet[row][col].status = 1;

                let mut stack:stack::StackLink = stack::StackLink::new();
                stack = topological_sort_from_cell(&sheet[row][col], sheet, stack);
                for cell in stack {
                    let r = cell.row;
                    let c = cell.col;
                    let mut res = 0;
                    match evaluate_expression(&cell.expression, rows, cols, sheet, &mut res, r, c, 0) {
                        0 | 1 => { sheet[r][c].borrow().val = res; sheet[r][c].borrow().status = 0; },
                        -2 => sheet[r][c].borrow().status = 1,
                        _ => {}
                    }
                }
                return -2;
            },
            -4 => return -4,
            -1 => return -1,
        }
    }

    -1
}


fn main() {
    let args: Vec<String> = env::args().collect();

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

    let start_time = SystemTime::now();

    if r < 1 || r > 999 {
        eprintln!("Invalid Input < 1<=R<=999 >");
        std::process::exit(-1);
    }

    if c < 1 || c > 18278 {
        eprintln!("Invalid Input < 1<=C<=18278 >");
        std::process::exit(-1);
    }

    let mut sheet: Vec<Vec<CellRef>> = vec![];
    create_sheet(&mut sheet);
    print_sheet(&sheet);

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

        let status = unsafe { execute_command(&input, R, C, &mut sheet) };

        if status == 1 {
            break;
        }

        let time_taken = start.elapsed().as_secs_f64();

        unsafe {
            if FLAG == 1 {
                print_sheet(&sheet);
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