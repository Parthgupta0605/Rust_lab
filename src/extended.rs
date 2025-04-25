use std::env;
use printpdf::{PdfDocument, PdfPage, PdfLayerIndex, BuiltinFont, Mm};
use crossterm::{
    cursor,
    event::{self, Event, KeyCode},
    style::{self, Color, SetForegroundColor},
    terminal::{self, ClearType},
    ExecutableCommand,
};
use std::collections::{HashMap, VecDeque, HashSet};
use std::fs::File;
use std::io::{self, stdout, BufReader, BufWriter, Write, Result};
use std::path::Path;
use serde::{Deserialize, Serialize};
use serde_json;

static mut START_ROW: usize = 0;
static mut START_COL: usize = 0;
static mut R :usize = 0;
static mut C :usize = 0;
// Cell struct to store data and metadata
#[derive(Clone, Debug, Serialize, Deserialize)]
struct Cell {
    raw_value: String,       // Raw input
    display_value: String,   // Value as displayed
    formula: Option<String>, // Formula if any
    is_locked: bool,         // Whether cell is locked
    alignment: Alignment,    // Text alignment
    width: usize,            // Cell width
    height: usize,           // Cell height
}

impl Cell {
    fn new() -> Self {
        Cell {
            raw_value: String::from("0"),
            display_value: String::from("0"),
            formula: None,
            is_locked: false,
            alignment: Alignment::Center,
            width: 5,  // Default width
            height: 1, // Default height
        }
    }

    fn display(&self) -> String {
        // Truncate if content exceeds cell width
        let content = if self.display_value.len() > self.width {
            self.display_value[0..self.width].to_string()
        } else {
            self.display_value.clone()
        };

        // Apply alignment
        match self.alignment {
            Alignment::Left => format!("{:<width$}", content, width = self.width),
            Alignment::Right => format!("{:>width$}", content, width = self.width),
            Alignment::Center => format!("{:^width$}", content, width = self.width),
        }
    }

    fn default() -> Self {
        Cell {
            raw_value: String::new(),
            display_value: String::new(),
            formula: None,
            alignment: Alignment::Center,
            is_locked: false,
            width: 5, // or whatever default width you use
            height: 1,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
enum Alignment {
    Left,
    Right,
    Center,
}

#[derive(Clone, Debug, PartialEq)]
enum Mode {
    Normal,
    Insert,
    Command,
    Find,
}

#[derive(Clone, Debug)]
struct CellAddress {
    col: usize,
    row: usize,
}

impl CellAddress {
    fn new(col: usize, row: usize) -> Self {
        CellAddress { col, row }
    }

    fn from_str(addr: &str) -> Option<Self> {
        if addr.len() < 2 {
            return None;
        }
        
        let col_char = addr.chars().next().unwrap();
        let col = match col_char {
            'A'..='Z' => (col_char as usize) - ('A' as usize),
            'a'..='z' => (col_char as usize) - ('a' as usize),
            _ => return None,
        };
        
        let row_str = &addr[1..];
        match row_str.parse::<usize>() {
            Ok(row) if row > 0 => Some(CellAddress::new(col, row - 1)),
            _ => None,
        }
    }

    fn col_to_letters(mut col: usize) -> String {
        let mut label = String::new();
        col += 1; // shift to 1-based
        while col > 0 {
            col -= 1;
            label.insert(0, (b'A' + (col % 26) as u8) as char);
            col /= 26;
        }
        label
    }
    
    fn to_string(&self) -> String {
       format!("{}{}", Self::col_to_letters(self.col), self.row + 1)
    }
}

#[derive(Clone, Debug)]
struct UndoAction {
    cell_address: CellAddress,
    old_cell: Cell,
}

struct SheetSnapshot {
    data: HashMap<String, Cell>,
    dependencies: HashMap<String, HashSet<String>>,
    dependents: HashMap<String, HashSet<String>>,
}

struct SheetAction {
    cells: Vec<UndoAction>,  // Collection of all cell changes in this action
}



struct Spreadsheet {
    data: HashMap<String, Cell>,
    cursor: CellAddress,
    mode: Mode,
    max_cols: usize,
    max_rows: usize,
    command_buffer: String,
    status_message: String,
    undo_stack: VecDeque<UndoAction>,
    redo_stack: VecDeque<UndoAction>,
    find_matches: Vec<CellAddress>,
    current_find_match: usize,
    find_query: String,
    dependents: HashMap<String, HashSet<String>>,  // Maps cell address to cells that depend on it
    dependencies: HashMap<String, HashSet<String>>,
    currently_updating: HashSet<String>, // Tracks cells being updated to prevent cycles
}

impl Spreadsheet {
    fn new(rows: usize, cols: usize) -> Self {
        let mut sheet = Spreadsheet {
            data: HashMap::new(),
            cursor: CellAddress::new(0, 0),
            mode: Mode::Normal,
            max_cols: cols,
            max_rows: rows,
            command_buffer: String::new(),
            status_message: String::new(),
            undo_stack: VecDeque::with_capacity(3),
            redo_stack: VecDeque::with_capacity(3),
            find_matches: Vec::new(),
            current_find_match: 0,
            find_query: String::new(),
            dependents: HashMap::new(),
            dependencies: HashMap::new(),
            currently_updating: HashSet::new(),
        };
        
        // Initialize cells
        for col in 0..cols {
            for row in 0..rows {
                let addr = CellAddress::new(col, row).to_string();
                sheet.data.insert(addr, Cell::new());
            }
        }
        
        sheet
    }

    fn get_cell(&self, addr: &CellAddress) -> Option<&Cell> {
        // println!("DEBUG: {}", addr.to_string());
        // println!("DEBUG: Current cell data: {:?}", self.data);
        self.data.get(&addr.to_string())
    }

    fn get_cell_mut(&mut self, addr: &CellAddress) -> Option<&mut Cell> {
        self.data.get_mut(&addr.to_string())
    }

    fn move_cursor(&mut self, dx: isize, dy: isize) {
        let new_col = self.cursor.col as isize + dx;
        let new_row = self.cursor.row as isize + dy;
        
        // Ensure within bounds
        if new_col >= 0 && new_col < self.max_cols as isize &&
           new_row >= 0 && new_row < self.max_rows as isize {
            self.cursor.col = new_col as usize;
            self.cursor.row = new_row as usize;
        }
    }

    fn jump_to_cell(&mut self, addr: &str) -> bool {
        if let Some(cell_addr) = CellAddress::from_str(addr) {
            if cell_addr.col < self.max_cols && cell_addr.row < self.max_rows {
                self.cursor = cell_addr;
                return true;
            }
        }
        false
    }

    fn add_dependency(&mut self, dependent: &str, dependency: &str) {
        // Record that 'dependent' depends on 'dependency'
        self.dependencies.entry(dependent.to_string())
            .or_insert_with(HashSet::new)
            .insert(dependency.to_string());
        
        // Record that 'dependency' is depended upon by 'dependent'
        self.dependents.entry(dependency.to_string())
            .or_insert_with(HashSet::new)
            .insert(dependent.to_string());

        println!("DEBUG: Added dependency: {} -> {}", dependent, dependency);
    }

    fn remove_dependencies(&mut self, cell_addr: &str) {
        // Remove all dependencies for this cell
        if let Some(deps) = self.dependencies.remove(cell_addr) {
            // For each dependency, remove this cell from its dependents
            for dep in deps {
                if let Some(dependents) = self.dependents.get_mut(&dep) {
                    dependents.remove(cell_addr);
                }
            }
        }
    }

    fn update_dependencies(&mut self, cell_addr: &str, formula: &str) {
        println!("DEBUG: Removing dependencies for cell {}", cell_addr);
        // First, remove any existing dependencies
        self.remove_dependencies(cell_addr);
        
        // Extract cell references from the formula
        // let dependencies = self.extract_dependencies(formula);
        
        // Add new dependencies
        if formula.starts_with('=') {

            let formula = &formula[1..]; // Skip the '=' character
            println!("DEBUG: Updating dependencies for formula {}", formula);
            // Handle range formulas like SUM(A1:B2)
            if formula.contains('(') && formula.contains(')') && formula.contains(':') {
                println!("DEBUG: Found range in formula");
                let range_start = formula.find('(').unwrap() + 1;
                let range_end = formula.find(')').unwrap();
                if range_start < range_end {
                    let range_str = &formula[range_start..range_end];
                    if let Some((start, end)) = self.parse_range(range_str) {
                        // Add all cells in the range as dependencies
                        for col in start.col..=end.col {
                            for row in start.row..=end.row {
                                let addr = CellAddress::new(col, row).to_string();
                                // dependencies.push(addr);
                                self.add_dependency(cell_addr, &addr);
                            }
                        }
                    }
                }
            } else if formula.contains('(') && formula.contains(')') {
                println!("DEBUG: Found function in formula");
                let func_start = formula.find('(').unwrap() + 1;
                let func_end = formula.find(')').unwrap();
                if func_start < func_end {
                    let cell_ref = &formula[func_start..func_end];
                    if let Some(addr) = CellAddress::from_str(cell_ref) {
                        // dependencies.push(addr.to_string());
                        self.add_dependency(cell_addr, &(addr.to_string()));
                    }
                }
            }
            // Handle simple cell references
            else {
                // Simple regex-like pattern for cell references (e.g., A1, B2)
                for c in formula.chars() {
                    if c.is_ascii_alphabetic() {
                        let col_char = c;
                        let mut remaining = formula.chars().skip_while(|&ch| ch != col_char).skip(1);
                        let mut row_str = String::new();
                        
                        while let Some(c) = remaining.next() {
                            if c.is_ascii_digit() {
                                row_str.push(c);
                            } else {
                                break;
                            }
                        }
                        
                        if !row_str.is_empty() {
                            if let Some(addr) = CellAddress::from_str(&format!("{}{}", col_char, row_str)) {
                                // dependencies.push(addr.to_string());
                                self.add_dependency(cell_addr, &(addr.to_string()));
                            }
                        }
                    }
                }
            }
        }
        
        // for dep in dependencies {
        //     self.add_dependency(cell_addr, &dep);
        // }
    }

    fn extract_dependencies(&self, formula: &str) -> Vec<String> {
        let mut dependencies = Vec::new();
        
        // Extract cell references from formulas like "=A1+B2"
        if formula.starts_with('=') {
            let formula = &formula[1..]; // Skip the '=' character
            
            // Handle range formulas like SUM(A1:B2)
            if formula.contains('(') && formula.contains(')') && formula.contains(':') {
                let range_start = formula.find('(').unwrap() + 1;
                let range_end = formula.find(')').unwrap();
                if range_start < range_end {
                    let range_str = &formula[range_start..range_end];
                    if let Some((start, end)) = self.parse_range(range_str) {
                        // Add all cells in the range as dependencies
                        for col in start.col..=end.col {
                            for row in start.row..=end.row {
                                let addr = CellAddress::new(col, row).to_string();
                                dependencies.push(addr);
                            }
                        }
                    }
                }
            } else if formula.contains('(') && formula.contains(')') {
                let func_start = formula.find('(').unwrap() + 1;
                let func_end = formula.find(')').unwrap();
                if func_start < func_end {
                    let cell_ref = &formula[func_start..func_end];
                    if let Some(addr) = CellAddress::from_str(cell_ref) {
                        dependencies.push(addr.to_string());
                    }
                }
            }
            // Handle simple cell references
            else {
                // Simple regex-like pattern for cell references (e.g., A1, B2)
                for c in formula.chars() {
                    if c.is_ascii_alphabetic() {
                        let col_char = c;
                        let mut remaining = formula.chars().skip_while(|&ch| ch != col_char).skip(1);
                        let mut row_str = String::new();
                        
                        while let Some(c) = remaining.next() {
                            if c.is_ascii_digit() {
                                row_str.push(c);
                            } else {
                                break;
                            }
                        }
                        
                        if !row_str.is_empty() {
                            if let Some(addr) = CellAddress::from_str(&format!("{}{}", col_char, row_str)) {
                                dependencies.push(addr.to_string());
                            }
                        }
                    }
                }
            }
        }
        
        dependencies
    }

    // fn propagate_changes(&mut self, cell_addr: &str) {
    //     // Find all cells that depend on this cell
    //     let dependents = if let Some(deps) = self.dependents.get(cell_addr) {
    //         deps.clone()
    //     } else {
    //         return;
    //     };
        
    //     // For each dependent, recalculate its value
    //     for dependent in dependents {
    //         if let Some(cell) = self.data.get(&dependent) {
    //             if let Some(formula) = &cell.formula {
    //                 // Clone the formula to avoid borrowing issues
    //                 let formula_clone = format!("={}", formula);
    //                 let addr = if let Some(addr) = CellAddress::from_str(&dependent) {
    //                     addr
    //                 } else {
    //                     continue;
    //                 };
                    
    //                 // Update the cell with its formula, which will recalculate its value
    //                 self.update_cell(&addr, &formula_clone);
                    
    //                 // Recursively propagate changes to cells that depend on this one
    //                 self.propagate_changes(&dependent);
    //             }
    //         }
    //     }
    // }
    fn propagate_changes(&mut self, cell_addr: &str) {
        // Get all cells that depend on this cell
        let mut dependents_to_process = Vec::new();
        
        // First, collect all the dependents without holding a reference to self
        if let Some(deps) = self.dependents.get(cell_addr) {
            for dep in deps {
                dependents_to_process.push(dep.clone());
            }
        } else {
            return;
        }
        println!("DEBUG: Dependents to process: {:?}", dependents_to_process);
        // Now process each dependent
        for dependent in dependents_to_process {
            // Check if the dependent is already being updated to avoid circular dependencies
            if self.currently_updating.contains(&dependent) {
                self.status_message = format!("ERROR: CIRCULAR DEPENDENCY DETECTED WITH {}", dependent);
                println!("DEBUG: Undo stack: {:?}", self.undo_stack);
                self.undo();
                self.status_message = format!("ERROR: CIRCULAR DEPENDENCY DETECTED WITH {}", dependent);
                return;
            }


            // Get the formula if it exists
            let formula_opt = if let Some(cell) = self.data.get(&dependent) {
                cell.formula.clone()
            } else {
                None
            };
            
            // If we have a formula, recalculate the cell
            if let Some(formula) = formula_opt {
                let formula_with_eq = format!("={}", formula);
                
                if let Some(addr) = CellAddress::from_str(&dependent) {
                    // Update the cell with its formula to recalculate
                    self.update_cell(&addr, &formula_with_eq, true);
                    
                    // We don't need to recursively call propagate_changes here
                    // because update_cell will handle that for us
                }
            }
        }
    }

    fn update_cell(&mut self, addr: &CellAddress, value: &str, multi:bool) -> bool {
        // First, check if cell exists and if it's locked
        let cell_exists = self.get_cell(addr).is_some();
        let is_locked = self.get_cell(addr).map_or(false, |cell| cell.is_locked);
        
        if !cell_exists {
            self.status_message = format!("ERROR: CELL {} NOT FOUND", addr.to_string());
            return false;
        }
        
        if is_locked {
            self.status_message = format!("ERROR: CELL {} LOCKED", addr.to_string());
            return false;
        }

        let cell_addr_str = addr.to_string();
        println!("DEBUG: Updating cell {} with value {}", cell_addr_str, value);
        println!("DEBUG: Currently updating: {:?}", self.currently_updating);
        // Check for circular dependency
        if self.currently_updating.contains(&cell_addr_str) {
            self.status_message = format!("ERROR: CIRCULAR DEPENDENCY DETECTED EARLY WITH {}", cell_addr_str);
            return false;
        }
        
        // Mark this cell as being updated
        self.currently_updating.insert(cell_addr_str.clone());

        // println!("Debug: Updating cell {} with value {}", addr.to_string(), value);
        // Save the old cell for undo (clone it before modifying)
        if let Some(old_cell) = self.get_cell(addr).cloned() {
            // Push to undo stack and clear redo stack
            // self.push_undo(addr.clone(), old_cell);
            // self.redo_stack.clear();    

            let mut is_valid_formula = false;
            if value.starts_with("=") {
                // Validate formula
                let formula = &value[1..];
                is_valid_formula = if formula.starts_with("SUM(") || formula.starts_with("MIN(") || formula.starts_with("MAX(") || formula.starts_with("STDEV(") {
                    if let Some(range_str) = formula.strip_prefix("SUM(").or_else(|| formula.strip_prefix("MIN("))
                        .or_else(|| formula.strip_prefix("MAX(")).or_else(|| formula.strip_prefix("STDEV("))
                        .and_then(|s| s.strip_suffix(')')) {
                        if let Some((start, end)) = self.parse_range(range_str) {
                            
                            let start_exists = self.get_cell(&start).is_some();
                            // println!("Debug: Start cell {} exists: {}", start.to_string(), start_exists);
                            let end_exists = self.get_cell(&end).is_some();
                            if(!(start_exists && end_exists)) {
                                self.status_message = format!("ERROR: INVALID RANGE {}", range_str);
                            }
                            start_exists && end_exists
                        } else {
                            self.status_message = format!("ERROR: INVALID RANGE {}", range_str);

                            false
                        }
                    } else {
                        self.status_message = format!("ERROR: INVALID RANGE {}", formula);
                        false
                    }
                } else if formula.starts_with("sqrt(") || formula.starts_with("log(") {
                    if let Some(arg) = formula.strip_prefix("sqrt(").or_else(|| formula.strip_prefix("log("))
                        .and_then(|s| s.strip_suffix(')')) {
                        CellAddress::from_str(arg).map_or(false, |addr| self.get_cell(&addr).is_some()) || arg.parse::<f64>().is_ok()
                    } else {
                        self.status_message = format!("ERROR: INVALID ARGUMENT {}", formula);
                        false
                    }
                } 
                else if formula.starts_with("(") && formula.ends_with(")") {
                    let cell_ref = &formula[1..formula.len() - 1];
                    if let Some(addr) = CellAddress::from_str(cell_ref) {
                        self.get_cell(&addr).is_some()
                    } else {
                        self.status_message = format!("ERROR: INVALID CELL REFERENCE {}", cell_ref);
                         false
                    }
                }
                else {
                    self.status_message = format!("ERROR: INVALID FORMULA {}", value);
                    false
                };
            }
            else {
                if(!multi){
                    println!("DEBUG: Pushing undo for cell {}", addr.to_string());
                    self.push_undo_sheet();
                    self.redo_stack.clear(); 
                }
                // self.push_undo_sheet();
                // self.redo_stack.clear(); 

                self.update_dependencies(&addr.to_string(), value);

                if let Some(mut cell) = self.get_cell_mut(addr) {
                    cell.formula = None;
                    cell.raw_value = value.to_string();
                    cell.display_value = value.to_string();
                }
                println!("DEBUG: propagating starting on {}", addr.to_string());

                self.propagate_changes(&addr.to_string());
                self.currently_updating.remove(&cell_addr_str);
        println!("DEBUG: Finished updating cell {}", cell_addr_str);
                return true;
            }
            if is_valid_formula {
                // Save the old cell for undo (clone it before modifying)
                if(!multi){
                    println!("DEBUG: Pushing undo for cell {}", addr.to_string());
                    self.push_undo_sheet();
                    self.redo_stack.clear(); 
                }

                let formula = &value[1..];
                // self.remove_dependencies(&addr.to_string());
                println!("DEBUG: Updating dependencies for cell {}", addr.to_string());
                self.update_dependencies(&addr.to_string(), value);
                // Compute the formula result
                let result = if formula.starts_with("SUM(") {
                    let range_str = formula.strip_prefix("SUM(").unwrap().strip_suffix(')').unwrap();
                    if let Some((start, end)) = self.parse_range(range_str) {
                        let mut sum = 0.0;
                        for col in start.col..=end.col {
                            for row in start.row..=end.row {
                                let addr = CellAddress::new(col, row);
                                if let Some(cell) = self.get_cell(&addr) {
                                    if let Ok(value) = cell.display_value.parse::<f64>() {
                                        sum += value;
                                    }
                                }
                            }
                        }
                        sum
                    } else {
                        0.0
                    }
                } else if formula.starts_with("MIN(") {
                    let range_str = formula.strip_prefix("MIN(").unwrap().strip_suffix(')').unwrap();
                    if let Some((start, end)) = self.parse_range(range_str) {
                        let mut min = f64::INFINITY;
                        for col in start.col..=end.col {
                            for row in start.row..=end.row {
                                let addr = CellAddress::new(col, row);
                                if let Some(cell) = self.get_cell(&addr) {
                                    if let Ok(value) = cell.display_value.parse::<f64>() {
                                        if value < min {
                                            min = value;
                                        }
                                    }
                                }
                            }
                        }
                        min
                    } else {
                        0.0
                    }
                } else if formula.starts_with("MAX(") {
                    let range_str = formula.strip_prefix("MAX(").unwrap().strip_suffix(')').unwrap();
                    if let Some((start, end)) = self.parse_range(range_str) {
                        let mut max = f64::NEG_INFINITY;
                        for col in start.col..=end.col {
                            for row in start.row..=end.row {
                                let addr = CellAddress::new(col, row);
                                if let Some(cell) = self.get_cell(&addr) {
                                    if let Ok(value) = cell.display_value.parse::<f64>() {
                                        if value > max {
                                            max = value;
                                        }
                                    }
                                }
                            }
                        }
                        max
                    } else {
                        0.0
                    }
                } else if formula.starts_with("STDEV(") {
                    let range_str = formula.strip_prefix("STDEV(").unwrap().strip_suffix(')').unwrap();
                    if let Some((start, end)) = self.parse_range(range_str) {
                        let mut values = Vec::new();
                        for col in start.col..=end.col {
                            for row in start.row..=end.row {
                                let addr = CellAddress::new(col, row);
                                if let Some(cell) = self.get_cell(&addr) {
                                    if let Ok(value) = cell.display_value.parse::<f64>() {
                                        values.push(value);
                                    }
                                }
                            }
                        }
                        let mean = values.iter().sum::<f64>() / values.len() as f64;
                        let variance = values.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / values.len() as f64;
                        variance.sqrt()
                    } else {
                        0.0
                    }
                } else if formula.starts_with("sqrt(") {
                    let arg = formula.strip_prefix("sqrt(").unwrap().strip_suffix(')').unwrap();
                    if let Ok(value) = arg.parse::<f64>() {
                        value.sqrt()
                    } else if let Some(addr) = CellAddress::from_str(arg) {
                        if let Some(cell) = self.get_cell(&addr) {
                            if let Ok(value) = cell.display_value.parse::<f64>() {
                                value.sqrt()
                            } else {
                                0.0
                            }
                        } else {
                            0.0
                        }
                    } else {
                        0.0
                    }
                } else if formula.starts_with("log(") {
                    let arg = formula.strip_prefix("log(").unwrap().strip_suffix(')').unwrap();
                    if let Ok(value) = arg.parse::<f64>() {
                        value.ln()
                    } else if let Some(addr) = CellAddress::from_str(arg) {
                        if let Some(cell) = self.get_cell(&addr) {
                            if let Ok(value) = cell.display_value.parse::<f64>() {
                                value.ln()
                            } else {
                                0.0
                            }
                        } else {
                            0.0
                        }
                    } else {
                        0.0
                    }
                } else if formula.starts_with("(") && formula.ends_with(")") {
                    println!("DEBUG: Found cell reference in formula");
                    let cell_ref = &formula[1..formula.len() - 1];
                    if let Some(addr) = CellAddress::from_str(cell_ref) {
                        if let Some(cell) = self.get_cell(&addr) {
                            if let Ok(value) = cell.display_value.parse::<f64>() {
                                value
                            } else {
                                0.0
                            }
                        } else {
                            0.0
                        }
                    } else {
                        0.0
                    }
                    
                }
                else {
                    0.0
                };
                // Update the cell's display value with the computed result
                if let Some(mut cell) = self.get_cell_mut(addr) {
                    cell.display_value = result.to_string();
                    cell.raw_value = result.to_string();
                    cell.formula = Some(value[1..].to_string());

                }
                println!("DEBUG: propagating starting on {}", addr.to_string());
                self.propagate_changes(&addr.to_string());
                self.currently_updating.remove(&cell_addr_str);
        println!("DEBUG: Finished updating cell {}", cell_addr_str);
                return true;
            }
            else {

                self.status_message = format!("ERROR: INVALID FORMULA {}", value);
                return false;
            }
        }
        // Ensure removal from currently_updating set in all cases
        
        return true;
    }
            
            // Now update the cell (we're done with operations that need to borrow self)
            
    //         if let Some(mut cell) = self.get_cell_mut(addr) {
    //             // Handle formula
    //             if value.starts_with("=") {
    //                 // Validate formula
    //                 let formula = &value[1..];
    //                 // let is_valid_formula = if formula.starts_with("SUM(") || formula.starts_with("MIN(") || formula.starts_with("MAX(") || formula.starts_with("STDEV(") {
    //                 //     if let Some(range_str) = formula.strip_prefix("SUM(").or_else(|| formula.strip_prefix("MIN("))
    //                 //         .or_else(|| formula.strip_prefix("MAX(")).or_else(|| formula.strip_prefix("STDEV("))
    //                 //         .and_then(|s| s.strip_suffix(')')) {
    //                 //         if let Some((start, end)) = self.parse_range(range_str) {
    //                 //             let start_exists = self.get_cell(&start).is_some();
    //                 //             let end_exists = self.get_cell(&end).is_some();
    //                 //             start_exists && end_exists
    //                 //         } else {
    //                 //             false
    //                 //         }
    //                 //     } else {
    //                 //         false
    //                 //     }
    //                 // } else if formula.starts_with("sqrt(") || formula.starts_with("log(") {
    //                 //     if let Some(arg) = formula.strip_prefix("sqrt(").or_else(|| formula.strip_prefix("log("))
    //                 //         .and_then(|s| s.strip_suffix(')')) {
    //                 //         CellAddress::from_str(arg).map_or(false, |addr| self.get_cell(&addr).is_some()) || arg.parse::<f64>().is_ok()
    //                 //     } else {
    //                 //         false
    //                 //     }
    //                 // } else {
    //                 //     false
    //                 // };

    //                 if !is_valid_formula {
    //                     self.status_message = format!("ERROR: INVALID FORMULA {}", value);
    //                     return false;
    //                 }
                        
    //                 cell.formula = Some(value[1..].to_string());
    //                 cell.raw_value = value.to_string();
    //                 // For now, just use formula as display value
    //                 cell.display_value = value.to_string();
    //             } else {
    //                 // println!("Debug :Updating cell {} with value {}", addr.to_string(), value);
                    
    //                 cell.formula = None;
    //                 cell.raw_value = value.to_string();
    //                 cell.display_value = value.to_string();

    //                 // println!("Debug: Cell {} updated to {}", addr.to_string(), cell.display_value);
    //             }
    //             return true;
    //         }
    //     }
        
    //     false
    // }
    // fn push_undo_sheet(&mut self) {
    //     // Create a copy of the entire sheet as individual cell actions
    //     let mut sheet_action = SheetAction {
    //         cells: Vec::new(),
    //     };
        
    //     // Add all cells to the action
    //     for (addr_str, cell) in &self.data {
    //         if let Some(addr) = CellAddress::from_str(addr_str) {
    //             sheet_action.cells.push(UndoAction {
    //                 cell_address: addr,
    //                 old_cell: cell.clone(),
    //             });
    //         }
    //     }
        
    //     // Maintain max 3 undo steps
    //     if self.undo_stack.len() >= 3 {
    //         // Remove oldest actions
    //         let actions_to_remove = sheet_action.cells.len();
    //         for _ in 0..actions_to_remove {
    //             if !self.undo_stack.is_empty() {
    //                 self.undo_stack.pop_front();
    //             }
    //         }
    //     }
        
    //     // Add all cells to the undo stack
    //     for cell_action in sheet_action.cells {
    //         self.undo_stack.push_back(cell_action);
    //     }
    // }

    fn push_undo(&mut self, addr: CellAddress, old_cell: Cell) {
        // Maintain max 3 undo steps
        if self.undo_stack.len() >= 3 {
            self.undo_stack.pop_front();
        }
        self.undo_stack.push_back(UndoAction {
            cell_address: addr,
            old_cell
        });
    }

    // fn undo(&mut self) -> bool {
    //     if let Some(action) = self.undo_stack.pop_back() {
    //         // Save current state for redo
    //         if let Some(cell) = self.get_cell(&action.cell_address) {
    //             // Push to redo stack
    //             self.redo_stack.push_back(UndoAction {
    //                 cell_address: action.cell_address.clone(),
    //                 old_cell: cell.clone()
    //             });
                
    //             // Apply the undo
    //             if let Some(target_cell) = self.get_cell_mut(&action.cell_address) {
    //                 *target_cell = action.old_cell;
    //                 self.status_message = "UNDO APPLIED".to_string();
    //                 return true;
    //             }
    //         }
    //     }
    //     self.status_message = "NOTHING TO UNDO".to_string();
    //     false
    // }

    // fn redo(&mut self) -> bool {
    //     if let Some(action) = self.redo_stack.pop_back() {
    //         // Save current state for undo
    //         if let Some(cell) = self.get_cell(&action.cell_address) {
    //             // Push to undo stack
    //             self.push_undo(action.cell_address.clone(), cell.clone());
                
    //             // Apply the redo
    //             if let Some(target_cell) = self.get_cell_mut(&action.cell_address) {
    //                 *target_cell = action.old_cell;
    //                 self.status_message = "REDO APPLIED".to_string();
    //                 return true;
    //             }
    //         }
    //     }
    //     self.status_message = "NOTHING TO REDO".to_string();
    //     false
    // }
    fn push_undo_sheet(&mut self) {
        // Add all cells to the undo stack
        for (addr_str, cell) in &self.data {
            if let Some(addr) = CellAddress::from_str(addr_str) {
                // Maintain max 3 undo steps - only check on the first cell
                if addr_str == "A1" && self.undo_stack.len() >= 3 {
                    self.undo_stack.clear();
                }
                
                self.undo_stack.push_back(UndoAction {
                    cell_address: addr,
                    old_cell: cell.clone(),
                });
            }
        }
    }

    fn undo(&mut self) -> bool {
        // Check if we have any actions to undo
        if self.undo_stack.is_empty() {
            self.status_message = "NOTHING TO UNDO".to_string();
            return false;
        }
        
        // Store all current cell states for redo before undoing
        for (addr_str, cell) in &self.data {
            if let Some(addr) = CellAddress::from_str(addr_str) {
                self.redo_stack.push_back(UndoAction {
                    cell_address: addr,
                    old_cell: cell.clone(),
                });
            }
        }
        
        // Now restore all cells from the undo stack
        let mut restored_cells = HashMap::new();
        
        while let Some(action) = self.undo_stack.pop_back() {
            // Store the restored cell
            restored_cells.insert(action.cell_address.to_string(), action.old_cell);
            
            // Stop when we've restored all cells
            if restored_cells.len() == self.data.len() {
                break;
            }
        }
        
        // Apply all restored cells to the sheet
        for (addr_str, cell) in restored_cells {
            if let Some(target_cell) = self.data.get_mut(&addr_str) {
                *target_cell = cell;
            }
        }
        
        self.status_message = "UNDO APPLIED".to_string();
        true
    }

    fn redo(&mut self) -> bool {
        // Check if we have any actions to redo
        if self.redo_stack.is_empty() {
            self.status_message = "NOTHING TO REDO".to_string();
            return false;
        }
        
        // Store all current cell states for undo before redoing
        for (addr_str, cell) in &self.data {
            if let Some(addr) = CellAddress::from_str(addr_str) {
                self.undo_stack.push_back(UndoAction {
                    cell_address: addr,
                    old_cell: cell.clone(),
                });
            }
        }
        
        // Now restore all cells from the redo stack
        let mut restored_cells = HashMap::new();
        
        while let Some(action) = self.redo_stack.pop_back() {
            // Store the restored cell
            restored_cells.insert(action.cell_address.to_string(), action.old_cell);
            
            // Stop when we've restored all cells
            if restored_cells.len() == self.data.len() {
                break;
            }
        }
        
        // Apply all restored cells to the sheet
        for (addr_str, cell) in restored_cells {
            if let Some(target_cell) = self.data.get_mut(&addr_str) {
                *target_cell = cell;
            }
        }
        
        self.status_message = "REDO APPLIED".to_string();
        true
    }

    fn recalculate_dependencies(&mut self) {
        // Clear existing dependencies
        self.dependencies.clear();
        self.dependents.clear();
        
        // Rebuild dependencies from formulas
        for (addr_str, cell) in &self.data {
            if let Some(formula) = &cell.formula {
                // Extract dependencies from formula and update the dependency maps
                let deps = self.extract_dependencies(formula);
                for dep in deps {
                    // Add this cell as dependent of the dependency
                    self.dependents.entry(dep.clone())
                        .or_insert_with(HashSet::new)
                        .insert(addr_str.clone());
                    
                    // Add the dependency to this cell's dependencies
                    self.dependencies.entry(addr_str.clone())
                        .or_insert_with(HashSet::new)
                        .insert(dep);
                }
            }
        }
    }

    fn lock_cell(&mut self, addr: Option<&str>) -> bool {
        let addr = if let Some(a) = addr {
            if let Some(cell_addr) = CellAddress::from_str(a) {
                cell_addr
            } else {
                return false;
            }
        } else {
            self.cursor.clone()
        };
        
        if let Some(cell) = self.get_cell_mut(&addr) {
            cell.is_locked = true;
            self.status_message = "CELL LOCKED".to_string();
            true
        } else {
            false
        }
    }

    fn unlock_cell(&mut self, addr: Option<&str>) -> bool {
        let addr = if let Some(a) = addr {
            if let Some(cell_addr) = CellAddress::from_str(a) {
                cell_addr
            } else {
                return false;
            }
        } else {
            self.cursor.clone()
        };
        
        if let Some(cell) = self.get_cell_mut(&addr) {
            cell.is_locked = false;
            self.status_message = "CELL UNLOCKED".to_string();
            true
        } else {
            false
        }
    }

    fn set_alignment(&mut self, addr: Option<&str>, align: &str) -> bool {
        let addr = if let Some(a) = addr {
            if let Some(cell_addr) = CellAddress::from_str(a) {
                cell_addr
            } else {
                return false;
            }
        } else {
            self.cursor.clone()
        };
        
        let alignment = match align {
            "l" => Alignment::Left,
            "r" => Alignment::Right,
            "c" => Alignment::Center,
            _ => return false,
        };
        
        if let Some(cell) = self.get_cell_mut(&addr) {
            if cell.is_locked {
                self.status_message = format!("ERROR: CELL {} LOCKED", addr.to_string());
                return false;
            }
            
            cell.alignment = alignment;
            self.status_message = "ALIGNMENT CHANGED".to_string();
            true
        } else {
            false
        }
    }

    fn set_dimension(&mut self, addr: Option<&str>, height: Option<usize>, width: Option<usize>) -> bool {
        println!("Debug: Setting dimension for cell {:?}", addr);
        let addr = if let Some(a) = addr {
            if let Some(cell_addr) = CellAddress::from_str(a) {
                cell_addr
            } else {
                return false;
            }
        } else {
            self.cursor.clone()
        };
        println!("Debug: Address after parsing: {:?}", addr);
        if let Some(cell) = self.get_cell_mut(&addr) {
            if cell.is_locked {
                self.status_message = format!("ERROR: CELL {} LOCKED", addr.to_string());
                return false;
            }
            println!("Debug: Cell found: {:?}", cell);
            if let Some(h) = height {
                println!("Debug: Setting height to {}", h);
                cell.height = h;
            }
            
            if let Some(w) = width {
                println!("Debug: Setting width to {}", w);
                cell.width = w;
            }
            
            self.status_message = "DIMENSION CHANGED".to_string();
            true
        } else {
            false
        }
    }

    fn find(&mut self, query: &str) -> bool {
        self.find_matches.clear();
        self.find_query = query.to_string();
        
        // Search for matches
        for col in 0..self.max_cols {
            for row in 0..self.max_rows {
                let addr = CellAddress::new(col, row);
                if let Some(cell) = self.get_cell(&addr) {
                    if cell.display_value.contains(query) {
                        self.find_matches.push(addr);
                    }
                }
            }
        }
        
        if !self.find_matches.is_empty() {
            self.current_find_match = 0;
            self.cursor = self.find_matches[0].clone();
            self.status_message = format!("{} MATCHES FOUND", self.find_matches.len());
            true
        } else {
            self.status_message = "NO MATCHES FOUND".to_string();
            false
        }
    }

    fn find_next(&mut self) -> bool {
        if self.find_matches.is_empty() {
            return false;
        }
        
        self.current_find_match = (self.current_find_match + 1) % self.find_matches.len();
        self.cursor = self.find_matches[self.current_find_match].clone();
        true
    }

    fn find_prev(&mut self) -> bool {
        if self.find_matches.is_empty() {
            return false;
        }
        
        if self.current_find_match == 0 {
            self.current_find_match = self.find_matches.len() - 1;
        } else {
            self.current_find_match -= 1;
        }
        
        self.cursor = self.find_matches[self.current_find_match].clone();
        true
    }

    fn parse_range(&self, range_str: &str) -> Option<(CellAddress, CellAddress)> {
        let parts: Vec<&str> = range_str.split(':').collect();
        if parts.len() != 2 {
            return None;
        }
        
        let start = CellAddress::from_str(parts[0])?;
        let end = CellAddress::from_str(parts[1])?;
        
        Some((start, end))
    }

    fn multi_insert(&mut self, range_str: &str, value: &str) -> bool {
        // Remove brackets if present
        let range_str = range_str.trim_start_matches('[').trim_end_matches(']');
        
        if let Some((start, end)) = self.parse_range(range_str) {
            let start_col = start.col.min(end.col);
            let end_col = start.col.max(end.col);
            let start_row = start.row.min(end.row);
            let end_row = start.row.max(end.row);
            self.push_undo_sheet();
            self.redo_stack.clear(); 
            for col in start_col..=end_col {
                for row in start_row..=end_row {
                    let addr = CellAddress::new(col, row);
                    if !self.update_cell(&addr, value,true) {
                        // If any cell fails (e.g., is locked), continue with the rest
                        continue;
                    }
                }
            }
            
            self.status_message = "MULTIPLE INSERTS".to_string();
            true
        } else {
            self.status_message = "INVALID RANGE".to_string();
            false
        }
    }

    fn save_json(&self, path: &Path) -> io::Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        serde_json::to_writer_pretty(writer, &self.data)?;
        Ok(())
    }

    // fn load_json(&mut self, path: &Path) -> io::Result<()> {
    //     let file = File::open(path)?;
    //     let reader = BufReader::new(file);
    //     self.data = serde_json::from_reader(reader)?;
    //     Ok(())
    // }

    fn load_json(&mut self, path: &Path) -> io::Result<()> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        self.data = serde_json::from_reader(reader)?;
        
        // Reset max rows and columns
        self.max_rows = 0;
        self.max_cols = 0;
        
        // Scan through all cell addresses to find the maximum row and column
        for addr_str in self.data.keys() {
            if let Some(addr) = CellAddress::from_str(addr_str) {
                // Update max_rows if this cell's row is larger
                if addr.row > self.max_rows {
                    self.max_rows = addr.row;
                }
                
                // Update max_cols if this cell's column is larger
                if addr.col > self.max_cols {
                    self.max_cols = addr.col;
                }
            }
        }
        
        // If no cells were found, set defaults
        if self.max_rows == 0 {
            self.max_rows = 10; // Default number of rows
        }
        
        if self.max_cols == 0 {
            self.max_cols = 10; // Default number of columns
        }
        self.max_rows += 1; // Adjust for 0-based indexing
        self.max_cols += 1; // Adjust for 0-based indexing
        // println!("DEBUG: Max rows: {}, Max cols: {}", self.max_rows, self.max_cols);
        unsafe {
            C = self.max_cols;
            R = self.max_rows;
        }
        
        Ok(())
    }

    // fn sort_range(&mut self, range_str: &str, ascending: bool) -> bool {
    //     // Remove brackets if present
    //     let range_str = range_str.trim_start_matches('[').trim_end_matches(']');
    
    //     if let Some((start, end)) = self.parse_range(range_str) {
    //         let col = start.col;
    //         let start_row = start.row;
    //         let end_row = end.row;
    
    //         // Collect full rows with the value in the sort column
    //         let mut rows: Vec<(usize, Vec<Cell>)> = Vec::new();
    
    //         for row in start_row..=end_row {
    //             let mut row_cells = Vec::new();
    //             for c in 0..unsafe{ C} {
    //                 let addr = CellAddress::new(c, row);
    //                 if let Some(cell) = self.get_cell(&addr).cloned() {
    //                     row_cells.push(cell);
    //                 } else {
    //                     row_cells.push(Cell::default()); // fallback empty cell
    //                 }
    //             }
    //             rows.push((row, row_cells));
    //         }
    
    //         // Sort rows based on value in the specified column
    //         rows.sort_by(|a, b| {
    //             let val_a = &a.1[col].display_value;
    //             let val_b = &b.1[col].display_value;
    //             let result = val_a.cmp(val_b);
    //             if ascending { result } else { result.reverse() }
    //         });
    
    //         // Apply sorted rows back
    //         for (i, (_, row_cells)) in rows.into_iter().enumerate() {
    //             let new_row = start_row + i;
    //             for (c, cell) in row_cells.into_iter().enumerate() {
    //                 let addr = CellAddress::new(c, new_row);
    //                 if let Some(target) = self.get_cell_mut(&addr) {
    //                     if !target.is_locked {
    //                         *target = cell;
    //                     }
    //                 }
    //             }
    //         }
    
    //         self.status_message = "ROW SORT APPLIED".to_string();
    //         true
    //     } else {
    //         self.status_message = "INVALID RANGE".to_string();
    //         false
    //     }
    // }
    fn sort_range(&mut self, range_str: &str, ascending: bool) -> bool {
        // Remove brackets if present
        let range_str = range_str.trim_start_matches('[').trim_end_matches(']');
    
        if let Some((start, end)) = self.parse_range(range_str) {
            let col = start.col;
            let start_row = start.row;
            let end_row = end.row;
    
            // Save the current state for undo before sorting
            self.push_undo_sheet();
            self.redo_stack.clear();
    
            // Collect full rows with the value in the sort column
            let mut rows: Vec<(usize, Vec<Cell>)> = Vec::new();
    
            for row in start_row..=end_row {
                let mut row_cells = Vec::new();
                for c in 0..self.max_cols {
                    let addr = CellAddress::new(c, row);
                    if let Some(cell) = self.get_cell(&addr).cloned() {
                        row_cells.push(cell);
                    } else {
                        row_cells.push(Cell::default()); // fallback empty cell
                    }
                }
                rows.push((row, row_cells));
            }
    
            // Sort rows based on value in the specified column
            rows.sort_by(|a, b| {
                let val_a = &a.1.get(col).map_or("", |cell| &cell.display_value);
                let val_b = &b.1.get(col).map_or("", |cell| &cell.display_value);
                
                // Try to compare as numbers first
                if let (Ok(num_a), Ok(num_b)) = (val_a.parse::<f64>(), val_b.parse::<f64>()) {
                    let result = num_a.partial_cmp(&num_b).unwrap_or(std::cmp::Ordering::Equal);
                    return if ascending { result } else { result.reverse() };
                }
                
                // If not numbers, compare as strings
                let result = val_a.cmp(val_b);
                if ascending { result } else { result.reverse() }
            });
    
            // Apply sorted rows back
            for (i, (_, row_cells)) in rows.into_iter().enumerate() {
                let new_row = start_row + i;
                for (c, cell) in row_cells.into_iter().enumerate() {
                    let addr = CellAddress::new(c, new_row);
                    if let Some(target) = self.get_cell_mut(&addr) {
                        if !target.is_locked {
                            *target = cell;
                        }
                    } else {
                        // Insert new cell if it doesn't exist
                        let addr_str = addr.to_string();
                        self.data.insert(addr_str, cell);
                    }
                }
            }
    
            self.status_message = "ROW SORT APPLIED".to_string();
            true
        } else {
            self.status_message = "INVALID RANGE".to_string();
            false
        }
    }

    fn format_cell_value(&self, addr: &CellAddress) -> String {
        let cell = self.get_cell(addr).clone().unwrap() else {
            return String::new(); // Return empty string if cell not found
        };
        let width = cell.width;
        let mut value = cell.display_value.clone();
        if value.len() > width {
            if width >= 3 {
                value = format!("{}..", &value[..width - 2]);
            } else {
                value = ".".repeat(width); // Not enough space for any content
            }
        }
        let padding = width.saturating_sub(value.len());
        
    
        match cell.alignment {
            Alignment::Left => format!("{:<width$}", value, width = width),
            Alignment::Right => format!("{:>width$}", value, width = width),
            Alignment::Center => {
                let left = padding / 2;
                let right = padding - left;
                format!(
                    "{}{}{}",
                    " ".repeat(left),
                    value,
                    " ".repeat(right)
                )
            }
        }
    }

    fn export_to_pdf(&self, filename: &str) -> Result<()> {
        // Create a new PDF document
        let (mut doc, page1, layer1) = PdfDocument::new("Spreadsheet Export", Mm(210.0), Mm(297.0), "Layer 1");
        let mut current_page = page1;
        let mut current_layer = doc.get_page(current_page).get_layer(layer1);
        
        // Add the built-in Helvetica font
        let font = doc.add_builtin_font(BuiltinFont::Helvetica).map_err(|e| {
            io::Error::new(io::ErrorKind::Other, format!("Error adding font: {}", e))
        })?;
        
        // Set page dimensions and layout parameters
        let page_width = Mm(210.0);  // A4 width
        let page_height = Mm(297.0); // A4 height
        let margin_top = Mm(20.0);
        let margin_bottom = Mm(20.0);
        let margin_left = Mm(10.0);
        let cell_width = Mm(19.0);   // Adjusted to fit 10 columns (A-J) plus row numbers
        let row_height = Mm(10.0);
        
        // Maximum rows per page calculation
        let content_height = page_height - margin_top - margin_bottom;
        let max_rows_per_page = (content_height.0 / row_height.0).floor() as i32 - 1; // -1 for header row
        
        // Calculate dimensions
        let row_count = unsafe { R };
        let col_count = unsafe { C };
        let max_cols = 10; // Limit to 10 columns (A-J)
        
        // Store page indices for adding page numbers later
        let mut page_indices = vec![page1];
        
        // Process the data in page chunks
        let mut processed_rows = 0;
        
        while processed_rows < row_count {
            // Calculate rows for current page
            let rows_in_this_page = std::cmp::min(max_rows_per_page,(row_count - processed_rows) as i32);
            let mut y_position = page_height - margin_top;
            
            // Draw column headers (A, B, C, etc.)
            let mut x_position = margin_left + cell_width; // Starting after row numbers column
            current_layer.use_text("", 10.0, margin_left, y_position, &font); // Empty top-left cell
            
            // Draw column headers A through J (limited to max_cols)
            for col in 0..std::cmp::min(col_count, max_cols) {
                let col_label = format!("{}", char::from(b'A' + col as u8));
                current_layer.use_text(&col_label, 10.0, x_position, y_position, &font);
                x_position += cell_width;
            }
            
            y_position -= row_height;
            
            // Draw rows with row numbers for this page
            for page_row in 0..rows_in_this_page {
                let actual_row = processed_rows + page_row as usize;
                
                // Draw row number
                let row_label = format!("{}", actual_row + 1); // +1 because row numbers start at 1
                current_layer.use_text(&row_label, 10.0, margin_left, y_position, &font);
                
                // Draw cells for this row
                x_position = margin_left + cell_width;
                for col in 0..std::cmp::min(col_count, max_cols) {
                    let addr = CellAddress::new(col, actual_row);
                    let text = if let Some(cell) = self.get_cell(&addr) {
                        cell.display_value.clone()
                    } else {
                        "".to_string()
                    };
                    
                    current_layer.use_text(&text, 10.0, x_position, y_position, &font);
                    x_position += cell_width;
                }
                
                y_position -= row_height;
            }
            
            processed_rows += rows_in_this_page as usize ;
            
            // Create a new page if there are more rows to process
            if processed_rows < row_count {
                let (new_page, new_layer) = doc.add_page(page_width, page_height, format!("Page {}", processed_rows / (max_rows_per_page as usize) + 2));
                current_page = new_page;
                current_layer = doc.get_page(current_page).get_layer(new_layer);
                page_indices.push(current_page); // Store the new page index
            }
        }
        
        // Add page numbers
        let page_count = page_indices.len();
        for (i, page_index) in page_indices.iter().enumerate() {
            let page_num = i + 1;
            let layer_ref = doc.get_page(*page_index).get_layer(layer1); // Reuse layer1 or create new layers
            
            // Add page number at bottom center
            let page_text = format!("Page {} of {}", page_num, page_count);
            layer_ref.use_text(&page_text, 10.0, page_width / 2.0 - Mm(15.0), margin_bottom / 2.0, &font);
        }
        
        // Save the document
        doc.save(&mut BufWriter::new(File::create(filename)?)).map_err(|e| {
            io::Error::new(io::ErrorKind::Other, format!("Error saving PDF: {}", e))
        })?;
        
        Ok(())
    }

    fn process_command(&mut self) -> bool {
        // First, copy the command buffer to a local String to avoid borrowing issues
        let cmd = self.command_buffer.trim().to_string();
        
        // Command parsing
        if cmd == "q" {
            return false; // Quit
        } else if cmd.starts_with("i") {
            // Enter insert mode
            self.mode = Mode::Insert;
            self.status_message = "INSERTING".to_string();
            
            // Check if a specific cell is specified
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts[0] != "i" {
                self.status_message = "INVALID COMMAND - Do you mean to write :i (cell name)".to_string();
            }
            if parts.len() > 1 {
                if !self.jump_to_cell(parts[1]) {
                    self.status_message = "INVALID CELL".to_string();
                }
            }
            self.command_buffer.clear(); // Clear command buffer before entering new value
        } else if cmd.starts_with("j") {
            // Jump to cell
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() > 1 {
                if !self.jump_to_cell(parts[1]) {
                    self.status_message = "INVALID CELL".to_string();
                }
            }
        } else if cmd == "undo" {
            self.undo();
        } else if cmd == "redo" {
            self.redo();
        } else if cmd.starts_with("find") {
            // Enter find mode
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() > 1 {
                if self.find(parts[1]) {
                    self.mode = Mode::Find;
                }
            } else {
                self.status_message = "INVALID FIND COMMAND".to_string();
            }
        } else if cmd.starts_with("mi") {
            // Multi-insert
            let parts: Vec<&str> = cmd.splitn(3, ' ').collect();
            if parts.len() == 3 {
                if !self.multi_insert(parts[1], parts[2]) {
                    self.status_message = "INVALID MULTI-INSERT".to_string();
                }
            } else {
                self.status_message = "INVALID MULTI-INSERT COMMAND".to_string();
            }
        } else if cmd.starts_with("lock") {
            // Lock cell
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() > 1 {
                if !self.lock_cell(Some(parts[1])) {
                    self.status_message = "INVALID LOCK COMMAND".to_string();
                }
            } else {
                self.lock_cell(None);
            }
        } else if cmd.starts_with("unlock") {
            // Unlock cell
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() > 1 {
                if !self.unlock_cell(Some(parts[1])) {
                    self.status_message = "INVALID UNLOCK COMMAND".to_string();
                }
            } else {
                self.unlock_cell(None);
            }
        } else if cmd.starts_with("align") || cmd.starts_with("allign") {
            // Set alignment
            let parts: Vec<&str> = cmd.splitn(3, ' ').collect();
            if parts.len() == 2 {
                // Just alignment for current cell
                if !self.set_alignment(None, parts[1]) {
                    self.status_message = "INVALID ALIGNMENT".to_string();
                }
            } else if parts.len() == 3 {
                // Cell and alignment
                if !self.set_alignment(Some(parts[1]), parts[2]) {
                    self.status_message = "INVALID ALIGNMENT COMMAND".to_string();
                }
            } else {
                self.status_message = "INVALID ALIGNMENT COMMAND".to_string();
            }
        } else if cmd.starts_with("dim") {
            // Set dimension
            // Format: :dim [cell] ((h,w))
            // Parse the command
            if cmd.contains('(') && cmd.contains(')') {
                let before_paren = cmd.split('(').next().unwrap().trim();
                let parts: Vec<&str> = before_paren.splitn(2, ' ').collect();
                
                // Parse the dimensions
                let dimension_str = cmd.split('(').nth(1).unwrap_or("").split(')').next().unwrap_or("");
                let dimensions: Vec<&str> = dimension_str.split(',').collect();
                
                let height = if dimensions.len() > 0 && !dimensions[0].trim().is_empty() {
                    dimensions[0].trim().parse::<usize>().ok()
                } else {
                    None
                };
                
                let width = if dimensions.len() > 1 && !dimensions[1].trim().is_empty() {
                    dimensions[1].trim().parse::<usize>().ok()
                } else {
                    None
                };
                println!("Debug: Height: {:?}, Width: {:?}", height, width);
                if parts.len() > 1 {
                    // Cell specified
                    if !self.set_dimension(Some(parts[1]), height, width) {
                        self.status_message = "INVALID DIMENSION COMMAND".to_string();
                    }
                } else {
                    // Current cell
                    if !self.set_dimension(None, height, width) {
                        self.status_message = "INVALID DIMENSION COMMAND".to_string();
                    }
                }
            } else {
                self.status_message = "INVALID DIMENSION FORMAT".to_string();
            }
        } else if cmd.starts_with("sort") {
            // Sort
            // Format: :sort [range] flag
            let parts: Vec<&str> = cmd.splitn(3, ' ').collect();
            if parts.len() == 3 {
                let ascending = parts[2] == "1";
                if !self.sort_range(parts[1], ascending) {
                    self.status_message = "INVALID SORT COMMAND".to_string();
                }
            } else {
                self.status_message = "INVALID SORT COMMAND".to_string();
            }
        } else if cmd.starts_with("saveas_") {
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() == 2 {
                let filetype = &cmd[7..cmd.find(' ').unwrap_or(cmd.len())];
                let filepath = parts[1].trim();
        
                match filetype {
                    "json" => {
                        if let Err(e) = self.save_json(Path::new(filepath)) {
                            self.status_message = format!("SAVE ERROR: {}", e);
                        } else {
                            self.status_message = format!("FILE SAVED TO {}", filepath);
                        }
                    }
                    "pdf" => {
                        if let Err(e) = self.export_to_pdf(filepath) {
                            self.status_message = format!("PDF EXPORT ERROR: {}", e);
                        } else {
                            self.status_message = format!("PDF SAVED TO {}", filepath);
                        }
                    }
                    _ => {
                        self.status_message = "UNSUPPORTED FORMAT. Use saveas_json or saveas_pdf.".to_string();
                    }
                }
            } else {
                self.status_message = "USAGE: saveas_<format> <filename>".to_string();
            }
        } else if cmd.starts_with("load") {
            // Load
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() == 2 {
                if let Err(e) = self.load_json(Path::new(parts[1])) {
                    self.status_message = format!("LOAD ERROR: {}", e);
                } else {
                    self.status_message = "FILE LOADED".to_string();
                }
            } else {
                self.status_message = "INVALID LOAD COMMAND".to_string();
            }
        } else if cmd == "hh" {
            // Go to leftmost cell in row
            self.cursor.col = 0;
        } else if cmd == "ll" {
            // Go to rightmost cell in row
            self.cursor.col = self.max_cols - 1;
        } else if cmd == "jj" {
            // Go to bottom cell in column
            self.cursor.row = self.max_rows - 1;
        } else if cmd == "kk" {
            // Go to top cell in column
            self.cursor.row = 0;
        } else {
            self.status_message = "INVALID COMMAND".to_string();
        }
        
        true // Continue running
    }

    fn handle_key_event(&mut self, key: KeyCode) -> bool {
        match self.mode {
            Mode::Normal => {
                match key {
                    KeyCode::Char('q') => return false, // Quit
                    KeyCode::Char('h') => self.move_cursor(-1, 0),
                    KeyCode::Char('j') => self.move_cursor(0, 1),
                    KeyCode::Char('k') => self.move_cursor(0, -1),
                    KeyCode::Char('l') => self.move_cursor(1, 0),
                    KeyCode::Char('w') => unsafe {
                        if START_ROW >= 10 {
                            START_ROW -= 10;
                        } else {
                            START_ROW = 0;
                        }
                    },
                    KeyCode::Char('d') => unsafe {
                        if START_COL + 20 <= unsafe {C} - 1 {
                            START_COL += 10;
                        } else {
                            START_COL = unsafe { C }.saturating_sub(10);
                        }
                    },
                    KeyCode::Char('a') => unsafe {
                        if START_COL >= 10 {
                            START_COL -= 10;
                        } else {
                            START_COL = 0;
                        }
                    },
                    KeyCode::Char('s') => unsafe {
                        if START_ROW + 20 <= unsafe {R} - 1 {
                            START_ROW += 10;
                        } else {
                            START_ROW = unsafe {R}.saturating_sub(10);
                        }
                    },
                    KeyCode::Char(':') => {
                        self.mode = Mode::Command;
                        self.command_buffer.clear();
                    },
                    _ => {}
                }
            },
            Mode::Insert => {
                match key {
                    KeyCode::Esc => {
                        self.mode = Mode::Normal;
                        self.status_message.clear();
                    },
                    KeyCode::Enter => {
                        // Apply changes and exit insert mode
                        // Clone the values to avoid borrowing issues
                        let cursor_clone = self.cursor.clone();
                        let command_buffer_clone = self.command_buffer.clone();
                        // println!("Debug: Inserting value {} at {}", command_buffer_clone, cursor_clone.to_string());
                        // Now we can safely call update_cell with the cloned values
                        self.status_message.clear();
                        self.update_cell(&cursor_clone, &command_buffer_clone, false);
                        self.mode = Mode::Normal;
                        self.command_buffer.clear();
                        
                    },
                    KeyCode::Backspace => {
                        self.command_buffer.pop();
                    },
                    KeyCode::Char(c) => {
                        self.command_buffer.push(c);
                    },
                    _ => {}
                }
            },
            Mode::Command => {
                match key {
                    KeyCode::Esc => {
                        self.mode = Mode::Normal;
                        self.command_buffer.clear();
                    },
                    KeyCode::Enter => {
                        self.mode = Mode::Normal;
                        let continue_running = self.process_command();
                        self.command_buffer.clear();
                        if !continue_running {
                            return false;
                        }
                    },
                    KeyCode::Backspace => {
                        self.command_buffer.pop();
                    },
                    KeyCode::Char(c) => {
                        self.command_buffer.push(c);
                    },
                    _ => {}
                }
            },
            Mode::Find => {
                match key {
                    KeyCode::Esc => {
                        self.mode = Mode::Normal;
                        self.find_matches.clear();
                        self.status_message.clear();
                    },
                    KeyCode::Char('n') => {
                        self.find_next();
                    },
                    KeyCode::Char('p') => {
                        self.find_prev();
                    },
                    _ => {}
                }
            }
        }
        
        true // Continue running
    }
    fn draw(&self, stdout: &mut io::Stdout) -> io::Result<()> {
        // Clear screen
        stdout.execute(terminal::Clear(ClearType::All))?;
        stdout.execute(cursor::MoveTo(0, 0))?;
        
        // Fixed widths for consistent display
        let row_label_width = 5;  // Width for row numbers column
        let cell_padding = 1;     // Space between cells
        let default_cell_width = 5; // Default width if no attribute specified
        
        // Calculate max width for each column based on cell's width attribute
        let mut col_widths = vec![default_cell_width; 10]; // Default width for each column
        
        // Calculate max width for each column
        for col in unsafe { START_COL..(START_COL + 10) } { 
            let col_idx = (col - unsafe { START_COL }) as usize;
            
            // Start with width needed for column header
            let col_letter = CellAddress::col_to_letters(col);
            col_widths[col_idx] = col_widths[col_idx].max(col_letter.len());
            
            // Check all cells in this column
            for row in unsafe { START_ROW..(START_ROW + 10).min(unsafe { R }) } {
                let addr = CellAddress::new(col, row);
                if let Some(cell) = self.get_cell(&addr) {
                    // Get cell width from attribute
                    col_widths[col_idx] = col_widths[col_idx].max(cell.width);
                }
            }
            
            // Ensure minimum width
            col_widths[col_idx] = col_widths[col_idx].max(3); // Minimum width of 3
        }
        
        // Draw header row with column labels
        stdout.execute(SetForegroundColor(Color::Cyan))?;
        // Empty space for the corner where row and column headers intersect
        write!(stdout, "{:<width$}", "", width = row_label_width+1)?;
        
        // Column headers (A, B, C, etc.)
        for col in unsafe { START_COL..(START_COL + 10).min(unsafe { C }) } {
            let col_idx = (col - unsafe { START_COL }) as usize;
            let col_letter = CellAddress::col_to_letters(col);
            let total_cell_width = col_widths[col_idx] + cell_padding;
            write!(stdout, "{:^width$}", col_letter, width = total_cell_width)?;
        }
    
        write!(stdout, "\r\n")?;
    
        
        // Draw grid rows
        for row in unsafe { START_ROW..(START_ROW + 10).min(unsafe { R }) } {
            // Row label - always in a fixed-width column
            stdout.execute(SetForegroundColor(Color::Cyan))?;
            write!(stdout, "{:>width$}", row + 1, width = row_label_width)?;
            stdout.execute(SetForegroundColor(Color::Reset))?;
            
            // Draw each cell in the row
            for col in unsafe {START_COL..(START_COL + 10).min(unsafe { C })} {
                let col_idx = (col - unsafe { START_COL }) as usize;
                let addr = CellAddress::new(col, row);
                let is_cursor_cell = col == self.cursor.col && row == self.cursor.row;
                
                // Add cell highlighting if this is the cursor position
                if is_cursor_cell {
                    stdout.execute(SetForegroundColor(Color::Black))?;
                    stdout.execute(style::SetBackgroundColor(Color::White))?;
                }
                
                // Display cell content with consistent spacing
                let mut cell_content = if let Some(cell) = self.get_cell(&addr) {
                    cell.display_value.clone()
                } else {
                    "0".to_string()
                };
                
                // Truncate content if it's too long for the column
                let available_width = col_widths[col_idx];
                if cell_content.len() > available_width {
                    cell_content = format!("{}..", &cell_content[0..available_width.saturating_sub(2)]);
                }
                // println!("Debug: addr {}", addr.to_string());
                write!(stdout, " {:^width$}", self.format_cell_value(&addr), width = col_widths[col_idx])?;
                
                // Reset styling after cell
                if is_cursor_cell {
                    stdout.execute(SetForegroundColor(Color::Reset))?;
                    stdout.execute(style::SetBackgroundColor(Color::Reset))?;
                }
            }
            
            write!(stdout, "\r\n")?;
        }
        
        // Rest of the function remains the same
        // Status bar
        writeln!(stdout)?;
        
        // Show current cell info
        if let Some(cell) = self.get_cell(&self.cursor) {
            let formula_text = match &cell.formula {
                Some(f) => f,
                None => "None",
            };
            
            let lock_status = if cell.is_locked { "Locked" } else { "Unlocked" };
            
            write!(stdout, "{} : {} | {} | {} ", 
                self.cursor.to_string(),
                cell.display_value,
                formula_text,
                lock_status
            )?;
        }
        
        // Display status message at the bottom right
        let (cols, rows) = terminal::size()?; // Get terminal size
        let status_message = &self.status_message;
        if !status_message.is_empty() {
            stdout.execute(cursor::MoveTo(cols.saturating_sub(status_message.len() as u16), rows.saturating_sub(1)))?;
            write!(stdout, "{}", status_message)?;
        }
        // Display command buffer at the bottom right
        if !self.command_buffer.is_empty() {
            let command_buffer = &self.command_buffer;
            stdout.execute(cursor::MoveTo(0, rows.saturating_sub(2)))?;
            write!(stdout, "{}", command_buffer)?;
        }
        
        stdout.flush()?;
        
        Ok(())
    }
}

pub fn run_extended() -> Result<()> {
    // Setup terminal

    let args: Vec<String> = env::args().collect();
    let (rows, cols) = if args.len() == 3 {
        let r = args[1].parse::<usize>().unwrap_or(10);
        let c = args[2].parse::<usize>().unwrap_or(10);
        (r, c)
    } else {
        eprintln!("Usage: {} <rows> <cols>. Defaulting to 10x10.", args[0]);
        (10, 10)
    };

    unsafe {
        R = rows;
        C = cols;
    }
    let mut stdout = stdout();
    terminal::enable_raw_mode()?;
    stdout.execute(terminal::Clear(ClearType::All))?;
    stdout.execute(cursor::Hide)?; // Hide cursor for custom rendering

    // Create spreadsheet (10x10 grid)
    let mut sheet = Spreadsheet::new(rows, cols);

    // Main event loop
    loop {
        // Draw the current state
        sheet.draw(&mut stdout)?;

        // Handle input
            // if event::poll(std::time::Duration::from_millis(100))? {
                if let Event::Key(key_event) = event::read()? {
                    if !sheet.handle_key_event(key_event.code) {
                        break; // Exit if handler returns false
                    }
                // }
            }
    }

    // Clean up
    terminal::disable_raw_mode()?;
    stdout.execute(cursor::Show)?; // Show cursor again
    stdout.execute(terminal::Clear(ClearType::All))?;
    stdout.execute(cursor::MoveTo(0, 0))?;
    
    println!("Thank you for using the Vim-like Spreadsheet!");
    Ok(())
}