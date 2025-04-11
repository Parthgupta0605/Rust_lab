use crate::cell::{push_dependent, pop_dependent, Cell, CellRef};
use crate::avl::{insert, delete_node, find};
use crate::stack::{push, StackNode};
use std::cell::RefCell;
use std::rc::Rc;
use std::time::SystemTime;

pub static mut FLAG: i32 = 1;
pub static mut R: usize = 0;
pub static mut C: usize = 0;
pub static mut START_ROW: usize = 0;
pub static mut START_COL: usize = 0;

pub fn create_sheet(r: usize, c: usize) -> Vec<Vec<CellRef>> {
    unsafe {
        R = r;
        C = c;
    }

    let mut sheet: Vec<Vec<CellRef>> = Vec::with_capacity(r);
    for _ in 0..r {
        let mut row: Vec<CellRef> = Vec::with_capacity(c);
        for _ in 0..c {
            row.push(Rc::new(RefCell::new(Cell::new())));
        }
        sheet.push(row);
    }
    sheet
}

pub fn add_dependency(c: &mut Cell, dep: CellRef, sheet: &mut Vec<Vec<CellRef>>) {
    unsafe {
        c.dependencies = insert(c.dependencies.take(), dep, sheet, C);
    }
}

pub fn add_dependent(c: &mut Cell, dep: CellRef) {
    push_dependent(c, dep);
}

pub fn delete_dependencies(cell1: &mut Cell, row: usize, col: usize, sheet: &mut Vec<Vec<CellRef>>) {
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
            let dep_cell = &node.cell;
            let dep_ptr = dep_cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
            let dep_row = dep_ptr / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
            let dep_col = dep_ptr / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
            if dfs(dep_cell, target, visited, dep_row, dep_col, sheet) {
                return true;
            }
            stack.push(node.left.clone());
            stack.push(node.right.clone());
        }
    }
    false
}

pub fn check_loop(
    start: &CellRef,
    target: &CellRef,
    start_row: usize,
    start_col: usize,
    target_row: usize,
    target_col: usize,
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
            let dep_cell = &node.cell;
            let dep_ptr = dep_cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
            let dep_row = dep_ptr / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
            let dep_col = dep_ptr / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };
            if dfs_range(dep_cell, visited, row1, col1, row2, col2, dep_row, dep_col, sheet) {
                return true;
            }
            stack.push(node.left.clone());
            stack.push(node.right.clone());
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

pub fn topological_sort_util(cell: &CellRef, visited: &mut Vec<bool>, sheet: &Vec<Vec<CellRef>>, stack: &mut Option<Box<StackNode>>) {
    let ptr_offset = cell.as_ptr() as usize - sheet[0][0].as_ptr() as usize;
    let row = ptr_offset / std::mem::size_of::<RefCell<Cell>>() / unsafe { C };
    let col = ptr_offset / std::mem::size_of::<RefCell<Cell>>() % unsafe { C };

    if !visited[row * unsafe { C } + col] {
        visited[row * unsafe { C } + col] = true;
        let cell_borrow = cell.borrow();
        let mut deps_stack = vec![cell_borrow.dependencies.clone()];
        while let Some(Some(node)) = deps_stack.pop() {
            topological_sort_util(&node.cell, visited, sheet, stack);
            deps_stack.push(node.left.clone());
            deps_stack.push(node.right.clone());
        }
        push(stack, Rc::clone(cell));
    }
}

pub fn topological_sort_from_cell(start_cell: &CellRef, sheet: &Vec<Vec<CellRef>>, stack: &mut Option<Box<StackNode>>) {
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
