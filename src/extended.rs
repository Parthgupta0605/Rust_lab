use std::env;
use crossterm::{
    cursor,
    event::{self, Event, KeyCode},
    style::{self, Color, SetForegroundColor},
    terminal::{self, ClearType},
    ExecutableCommand,
};
use std::collections::{HashMap, VecDeque};
use std::fs::File;
use std::io::{self, stdout, BufReader, BufWriter, Write, Result};
use std::path::Path;
use serde::{Deserialize, Serialize};
use serde_json;

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
            alignment: Alignment::Left,
            width: 8,  // Default width
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

    fn to_string(&self) -> String {
        format!("{}{}", 
                ((self.col as u8) + b'A') as char, 
                self.row + 1)
    }
}

#[derive(Clone, Debug)]
struct UndoAction {
    cell_address: CellAddress,
    old_cell: Cell,
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
}

impl Spreadsheet {
    fn new(cols: usize, rows: usize) -> Self {
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

    fn update_cell(&mut self, addr: &CellAddress, value: &str) -> bool {
        // First, check if cell exists and if it's locked
        let cell_exists = self.get_cell(addr).is_some();
        let is_locked = self.get_cell(addr).map_or(false, |cell| cell.is_locked);
        
        if !cell_exists {
            return false;
        }
        
        if is_locked {
            self.status_message = format!("ERROR: CELL {} LOCKED", addr.to_string());
            return false;
        }
        
        // Save the old cell for undo (clone it before modifying)
        if let Some(old_cell) = self.get_cell(addr).cloned() {
            // Push to undo stack and clear redo stack
            self.push_undo(addr.clone(), old_cell);
            self.redo_stack.clear();
            
            // Now update the cell (we're done with operations that need to borrow self)
            let cell_clone = self.get_cell(addr).cloned();
            if let Some(mut cell) = cell_clone {
                // Handle formula
                if value.starts_with("=") {
                    // Validate formula
                    let formula = &value[1..];
                    let is_valid_formula = if formula.starts_with("SUM(") || formula.starts_with("MIN(") || formula.starts_with("MAX(") || formula.starts_with("STDEV(") {
                        if let Some(range_str) = formula.strip_prefix("SUM(").or_else(|| formula.strip_prefix("MIN("))
                            .or_else(|| formula.strip_prefix("MAX(")).or_else(|| formula.strip_prefix("STDEV("))
                            .and_then(|s| s.strip_suffix(')')) {
                            if let Some((start, end)) = self.parse_range(range_str) {
                                let start_exists = self.get_cell(&start).is_some();
                                let end_exists = self.get_cell(&end).is_some();
                                start_exists && end_exists
                            } else {
                                false
                            }
                        } else {
                            false
                        }
                    } else if formula.starts_with("sqrt(") || formula.starts_with("log(") {
                        if let Some(arg) = formula.strip_prefix("sqrt(").or_else(|| formula.strip_prefix("log("))
                            .and_then(|s| s.strip_suffix(')')) {
                            CellAddress::from_str(arg).map_or(false, |addr| self.get_cell(&addr).is_some()) || arg.parse::<f64>().is_ok()
                        } else {
                            false
                        }
                    } else {
                        false
                    };

                    if !is_valid_formula {
                        self.status_message = format!("ERROR: INVALID FORMULA {}", value);
                        return false;
                    }
                    cell.formula = Some(value[1..].to_string());
                    cell.raw_value = value.to_string();
                    // For now, just use formula as display value
                    cell.display_value = value.to_string();
                } else {
                    cell.formula = None;
                    cell.raw_value = value.to_string();
                    cell.display_value = value.to_string();
                }
                return true;
            }
        }
        
        false
    }

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

    fn undo(&mut self) -> bool {
        if let Some(action) = self.undo_stack.pop_back() {
            // Save current state for redo
            if let Some(cell) = self.get_cell(&action.cell_address) {
                // Push to redo stack
                self.redo_stack.push_back(UndoAction {
                    cell_address: action.cell_address.clone(),
                    old_cell: cell.clone()
                });
                
                // Apply the undo
                if let Some(target_cell) = self.get_cell_mut(&action.cell_address) {
                    *target_cell = action.old_cell;
                    self.status_message = "UNDO APPLIED".to_string();
                    return true;
                }
            }
        }
        self.status_message = "NOTHING TO UNDO".to_string();
        false
    }

    fn redo(&mut self) -> bool {
        if let Some(action) = self.redo_stack.pop_back() {
            // Save current state for undo
            if let Some(cell) = self.get_cell(&action.cell_address) {
                // Push to undo stack
                self.push_undo(action.cell_address.clone(), cell.clone());
                
                // Apply the redo
                if let Some(target_cell) = self.get_cell_mut(&action.cell_address) {
                    *target_cell = action.old_cell;
                    self.status_message = "REDO APPLIED".to_string();
                    return true;
                }
            }
        }
        self.status_message = "NOTHING TO REDO".to_string();
        false
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
            if cell.is_locked {
                self.status_message = format!("ERROR: CELL {} LOCKED", addr.to_string());
                return false;
            }
            
            if let Some(h) = height {
                cell.height = h;
            }
            
            if let Some(w) = width {
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
            
            for col in start_col..=end_col {
                for row in start_row..=end_row {
                    let addr = CellAddress::new(col, row);
                    if !self.update_cell(&addr, value) {
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

    fn load_json(&mut self, path: &Path) -> io::Result<()> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        self.data = serde_json::from_reader(reader)?;
        Ok(())
    }

    fn sort_range(&mut self, range_str: &str, ascending: bool) -> bool {
        // Remove brackets if present
        let range_str = range_str.trim_start_matches('[').trim_end_matches(']');
        
        if let Some((start, end)) = self.parse_range(range_str) {
            // Currently only supporting sorting a single column
            if start.col == end.col {
                let col = start.col;
                let start_row = start.row.min(end.row);
                let end_row = start.row.max(end.row);
                
                // Collect values to sort
                let mut values: Vec<(usize, String)> = Vec::new();
                for row in start_row..=end_row {
                    let addr = CellAddress::new(col, row);
                    if let Some(cell) = self.get_cell(&addr) {
                        values.push((row, cell.display_value.clone()));
                    }
                }
                
                // Sort values
                values.sort_by(|a, b| {
                    let result = a.1.cmp(&b.1);
                    if ascending { result } else { result.reverse() }
                });
                
                // Apply sorted values (this is a simplified approach)
                for (i, (row, _value)) in values.iter().enumerate() {
                    let source_addr = CellAddress::new(col, *row);
                    let target_addr = CellAddress::new(col, start_row + i);
                    
                    if let Some(source_cell) = self.get_cell(&source_addr).cloned() {
                        if let Some(target_cell) = self.get_cell_mut(&target_addr) {
                            if !target_cell.is_locked {
                                *target_cell = source_cell;
                            }
                        }
                    }
                }
                
                self.status_message = "SORT APPLIED".to_string();
                true
            } else {
                self.status_message = "ONLY COLUMN SORTING IMPLEMENTED".to_string();
                false
            }
        } else {
            self.status_message = "INVALID RANGE".to_string();
            false
        }
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
        } else if cmd.starts_with("save") {
            // Save
            let parts: Vec<&str> = cmd.splitn(2, ' ').collect();
            if parts.len() == 2 {
                if let Err(e) = self.save_json(Path::new(parts[1])) {
                    self.status_message = format!("SAVE ERROR: {}", e);
                } else {
                    self.status_message = "FILE SAVED".to_string();
                }
            } else {
                self.status_message = "INVALID SAVE COMMAND".to_string();
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
                        
                        // Now we can safely call update_cell with the cloned values
                        self.update_cell(&cursor_clone, &command_buffer_clone);
                        self.mode = Mode::Normal;
                        self.command_buffer.clear();
                        self.status_message.clear();
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
        let cell_width = 5;       // Width for each cell
        let cell_padding = 1;     // Space between cells
        let total_cell_width = cell_width + cell_padding;
        
        // Draw header row with column labels
        stdout.execute(SetForegroundColor(Color::Cyan))?;
        // Empty space for the corner where row and column headers intersect
        write!(stdout, "{:<width$}", "", width = row_label_width+1)?;
        
        // Column headers (A, B, C, etc.)
        for col in 0..self.max_cols {
            let col_letter = (b'A' + col as u8) as char;
            write!(stdout, "{:^width$}", col_letter, width = total_cell_width)?;
        }

        write!(stdout, "\r\n")?;

        
        // Draw grid rows
        for row in 0..self.max_rows {
            // Row label - always in a fixed-width column
            stdout.execute(SetForegroundColor(Color::Cyan))?;
            write!(stdout, "{:>width$}", row + 1, width = row_label_width)?;
            stdout.execute(SetForegroundColor(Color::Reset))?;
            
            // Draw each cell in the row
            for col in 0..self.max_cols {
                let addr = CellAddress::new(col, row);
                let is_cursor_cell = col == self.cursor.col && row == self.cursor.row;
                
                // Add cell highlighting if this is the cursor position
                if is_cursor_cell {
                    stdout.execute(SetForegroundColor(Color::Black))?;
                    stdout.execute(style::SetBackgroundColor(Color::White))?;
                }
                
                // Display cell content with consistent spacing
                let cell_content = if let Some(cell) = self.get_cell(&addr) {
                    cell.display_value.clone()
                } else {
                    "0".to_string()
                };
                
                write!(stdout, " {:^width$}", cell_content, width = cell_width)?;
                
                // Reset styling after cell
                if is_cursor_cell {
                    stdout.execute(SetForegroundColor(Color::Reset))?;
                    stdout.execute(style::SetBackgroundColor(Color::Reset))?;
                }

            }
            // stdout.execute(cursor::MoveTo((row+1).try_into().unwrap(),0))?;
            write!(stdout, "\r\n")?;

        }
        
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

        // Rest of the status display code remains unchanged
        stdout.flush()?;
        
        Ok(())
    }
}

pub fn run_extended() -> Result<()> {
    // Setup terminal
    let args: Vec<String> = env::args().collect();
    let (rows, cols) = if args.len() >= 4 {  // Changed from 3 to 4 to account for the -vim flag
        let r = args[2].parse::<usize>().unwrap_or(10);
        let c = args[3].parse::<usize>().unwrap_or(10);
        (r, c)
    } else {
        eprintln!("Usage: {} -vim <rows> <cols>. Defaulting to 10x10.", args[0]);
        (10, 10)
    };
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