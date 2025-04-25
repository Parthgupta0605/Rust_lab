//! # Extended Spreadsheet Module for Hacker Spreadsheet
//!
//! This module extends the basic functionality of the spreadsheet program by 
//! implementing a text-based editor inspired by Vim, specifically designed for
//! terminal users. The extension aims to enhance the usability and functionality 
//! of the original spreadsheet program, allowing for a keyboard-driven, privacy-focused 
//! experience with remote editing capabilities.
use std::env;
use printpdf::{PdfDocument,  BuiltinFont, Mm};
use crossterm::{
    cursor::{MoveTo,Show,Hide,position},
    event::{self, Event, KeyCode},
    style::{self, Color, SetForegroundColor},
    terminal::{self,Clear, ClearType},
    ExecutableCommand,
};
use std::collections::{HashMap, VecDeque, HashSet};
use std::fs::File;
use std::io::{self, stdout, BufReader, BufWriter, Write, Result};
use std::path::Path;
use serde::{Deserialize, Serialize};
use serde_json;
use std::process::{ Stdio};
use rand::seq::SliceRandom;
use std::thread;



use rodio::{OutputStream, Sink};
use std::time::{Duration, Instant};

/// A static mutable variable to store the starting row for displaying the spreadsheet. 
static mut START_ROW: usize = 0;
/// A static mutable variable to store the starting column for displaying the spreadsheet.
static mut START_COL: usize = 0;
/// A static mutable variable to store the number of rows in the spreadsheet.
static mut R :usize = 0;
/// A static mutable variable to store the number of columns in the spreadsheet.
static mut C :usize = 0;


/// Plays a sound synchronously using Windows PowerShell.
///
/// This function takes a file path to a `.wav` sound file and uses PowerShell to play it
/// via the `Media.SoundPlayer` class. The playback is synchronous, meaning the function
/// waits until the sound finishes before returning.
///
/// # Arguments
/// * `path` - A string slice representing the file path to the `.wav` file. Forward slashes
///   will be automatically converted to backslashes for Windows compatibility.
///
/// # Platform
/// This function is intended for Windows systems and requires PowerShell to be available
/// in the system path.
///
/// # Panics
/// Panics if PowerShell fails to launch or execute the command.
///
/// # Example
/// ```
/// play_sound("C:/path/to/sound.wav");
/// ```

pub fn play_sound(path: &str) {
    // Convert path to Windows-style and launch PowerShell
    let win_path = path.replace("/", "\\");
    let command = format!("(New-Object Media.SoundPlayer '{}').PlaySync();", win_path);

    std::process::Command::new("powershell.exe")
        .args(["-c", &command])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .expect("Failed to play sound via PowerShell");
}

/// Triggers a visual and audio-based jump scare in the terminal.
///
/// This function is part of the Haunt Mode experience. It performs the following actions:
/// - Plays a predefined scream sound effect to startle the user.
/// - Clears the terminal and displays a centered, red-colored ASCII art scare message.
/// - Temporarily halts execution to allow the user to experience the full effect.
///
/// The ASCII art is centered based on the current terminal size to maximize visual impact.
/// Intended for brief use during themed or playful modes of the application.
///
/// # Notes
/// - Sound playback depends on the availability and correctness of the `play_sound` function and audio file path.
/// - The function requires terminal control capabilities provided by the `crossterm` crate.

fn trigger_jump_scare() {

    let mut stdout = stdout();

    // 🧨 Play scream sound
    let scream_path = r#"C:\Users\hp\OneDrive - IIT Delhi\Desktop\Academics\prisha_rust_lab\scary-scream.wav"#; // Path to your sound file
    play_sound(scream_path);
    let scare_art = r#"
    ████████████████████████████████████████
    █                                      █
    █           👻  𝓑𝓞𝓞𝓞𝓞𝓞!               █
    █      it has already begun            █
    █     ░░░ the veil is thinning ░░░     █
    █         you were warned...           █
    █                                      █
    ████████████████████████████████████████
    "#;

    // Optional: center it on screen
    let (cols, rows) = crossterm::terminal::size().unwrap();
    let x = cols.saturating_sub(40) / 2;
    let y = rows.saturating_sub(6) / 2;

    stdout.execute(Clear(ClearType::All)).unwrap();
    stdout.execute(MoveTo(x, y)).unwrap();
    stdout.execute(SetForegroundColor(Color::Red)).unwrap();

    for line in scare_art.lines() {
        let (x, y) = position().unwrap();
        stdout.execute(MoveTo(x, y)).unwrap();
        writeln!(stdout, "{}", line).unwrap();
    }

    stdout.flush().unwrap();
    thread::sleep(Duration::from_secs(2));
}

// Cell struct to store data and metadata
/// Represents a single cell in the spreadsheet.
///
/// The `Cell` struct holds both the raw input value (as entered by the user) and the 
/// value to be displayed in the spreadsheet. It also supports formulas, text alignment, 
/// and cell dimensions (width and height). The cell can be locked to prevent editing.
///
/// # Fields:
/// - `raw_value`: The raw input string (e.g., numbers, text, or formulas).
/// - `display_value`: The value that will be shown to the user, possibly altered by formulas.
/// - `formula`: An optional string containing a formula that is applied to compute the value.
/// - `is_locked`: A boolean indicating whether the cell is locked and cannot be edited.
/// - `alignment`: The alignment of the text inside the cell (e.g., left, right, or center).
/// - `width`: The width of the cell (in characters).
/// - `height`: The height of the cell (in rows).
/// # Methods:
/// - `new`: Creates a new `Cell` with default values.
/// - `display`: Returns the content of the cell formatted according to its alignment and width.
/// - `default`: Creates a new, default `Cell` with empty values for `raw_value` and `display_value`.
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
/// Represents the alignment of text within a cell.
///
/// The `Alignment` enum defines the available text alignments for a cell:
/// - `Left`: Aligns text to the left side of the cell.
/// - `Right`: Aligns text to the right side of the cell.
/// - `Center`: Centers the text in the middle of the cell.
#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
enum Alignment {
    Left,
    Right,
    Center,
}
/// Represents different modes the spreadsheet can be in.
///
/// The `Mode` enum defines the available modes for the spreadsheet editor:
/// - `Normal`: Default mode for interacting with the spreadsheet.
/// - `Insert`: Mode for inserting new data or formulas into cells.
/// - `Command`: Mode for executing commands.
/// - `Find`: Mode for searching within the spreadsheet.
#[derive(Clone, Debug, PartialEq)]
enum Mode {
    Normal,
    Insert,
    Command,
    Find,
}
/// Represents a cell's address in the spreadsheet using column and row indices.
///
/// The `CellAddress` struct holds the `col` (column index) and `row` (row index) for a specific
/// cell, and provides methods for converting between string representations of cell addresses
/// (e.g., "A1", "B2") and the internal column/row index format.
///
/// # Methods:
/// - `new`: Creates a new `CellAddress` from a column and row index.
/// - `from_str`: Parses a string (e.g., "A1", "B2") into a `CellAddress` if valid.
/// - `col_to_letters`: Converts a column index to the corresponding Excel-style column label (e.g., 0 -> "A", 1 -> "B", 26 -> "AA").
#[derive(Clone, Debug)]
struct CellAddress {
    col: usize,
    row: usize,
}

impl CellAddress {
    /// Creates a new `CellAddress` from a column and row index.
    ///
    /// # Arguments:
    /// - `col`: The zero-based column index (0 for 'A').
    /// - `row`: The zero-based row index (0 for row 1).
    ///
    /// # Returns:
    /// A `CellAddress` struct representing the cell at the specified position.
    fn new(col: usize, row: usize) -> Self {
        CellAddress { col, row }
    }
    /// Parses a string (e.g., "A1", "B2") into a `CellAddress`.
    ///
    /// The string must be in the format of a letter (column) followed by a number (row),
    /// such as "A1" or "B2". The column is case-insensitive.
    ///
    /// # Arguments:
    /// - `addr`: A string representing the cell address, e.g., "A1", "B2".
    ///
    /// # Returns:
    /// An `Option<CellAddress>`, which is `Some(CellAddress)` if the string is valid,
    /// or `None` if the string is invalid.
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
    /// Converts a column index to an Excel-style column label (e.g., 0 -> "A", 1 -> "B", 26 -> "AA").
    ///
    /// # Arguments:
    /// - `col`: The zero-based column index.
    ///
    /// # Returns:
    /// A string representing the Excel-style column label.
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
      /// Converts the `CellAddress` to a string representation (e.g., "A1", "B2").
    ///
    /// # Returns:
    /// A string representing the cell address in the format "A1", "B2", etc.
    fn to_string(&self) -> String {
       format!("{}{}", Self::col_to_letters(self.col), self.row + 1)
    }
}

// Represents an undo action in the spreadsheet, storing the state of a cell before an edit.
///
/// The `UndoAction` struct holds information about a cell's address and its previous state (the `old_cell`),
/// allowing for the undoing of a specific change made to a cell. This can be useful for implementing 
/// undo functionality in the spreadsheet editor.
///
/// # Fields:
/// - `cell_address`: The address of the cell that was modified.
/// - `old_cell`: The previous state of the cell before the edit was made, including its value, formula, and other properties.

#[derive(Clone, Debug)]
struct UndoAction {
    cell_address: CellAddress,
    old_cell: Cell,
}

// Represents a collection of cell changes in a single action that can be undone or redone.
//
// The `SheetAction` struct groups multiple `UndoAction` instances that represent the changes made to cells
// during a particular operation. This structure is useful for tracking the state of a spreadsheet during edits
// and facilitates undo and redo functionality.
//
// # Fields:
// - `cells`: A collection of all `UndoAction` instances, representing the changes made to individual cells
//   in the current action.
// struct SheetAction {
//     cells: Vec<UndoAction>,  // Collection of all cell changes in this action
// }


/// Represents the state of the entire spreadsheet, including cell data, user interaction, and tracking of undo/redo actions.
///
/// The `Spreadsheet` struct encapsulates the entire state of a spreadsheet, including the data of each cell,
/// the current cursor position, the mode of operation (e.g., normal, insert), and additional attributes to manage
/// user actions such as undo, redo, and search. It also manages dependencies between cells and tracks changes
/// in real-time to ensure consistent updates across the spreadsheet.
///
/// # Fields:
/// - `data`: A `HashMap` storing the actual data (cells) of the spreadsheet, where the key is the cell address.
/// - `cursor`: The current position of the cursor (cell address).
/// - `mode`: The current mode of the spreadsheet (e.g., Normal, Insert, Command, Find).
/// - `max_cols`: The maximum number of columns in the spreadsheet.
/// - `max_rows`: The maximum number of rows in the spreadsheet.
/// - `command_buffer`: A string buffer for storing the current command being entered by the user.
/// - `status_message`: A message that displays the current status or feedback for the user.
/// - `undo_stack`: A stack (using `VecDeque`) that tracks the history of actions that can be undone.
/// - `redo_stack`: A stack (using `VecDeque`) that tracks the history of undone actions that can be redone.
/// - `find_matches`: A list of `CellAddress` instances that match the current search query.
/// - `current_find_match`: The index of the current match in the `find_matches` list.
/// - `find_query`: The current search query being used to find matches in the spreadsheet.
/// - `dependents`: A `HashMap` mapping a cell address to the set of cells that depend on it.
/// - `dependencies`: A `HashMap` mapping a cell address to the set of cells it depends on.
/// - `currently_updating`: A set of cell addresses currently being updated, used to avoid cycles in dependency resolution.
/// ### Haunt Mode & Visual Effects:
/// - `haunted`: Indicates whether Haunt Mode is active.
/// - `haunt_sink`: Optional `Sink` for playing haunted audio effects.
/// - `haunt_stream`: Optional `OutputStream` tied to the haunted audio.
/// - `flicker_on`: Enables screen flicker effects when Haunt Mode is active.
/// - `last_flicker`: Timestamp of the last flicker event, used to control flicker intervals.
/// - `corruption_level`: Represents the current level of screen corruption (0–3).
/// - `last_corruption_tick`: Timestamp of the last corruption update.
/// - `haunted_start`: Records when Haunt Mode was activated.
/// - `jump_scare_triggered`: Tracks whether a jump scare has already occurred during Haunt Mode.
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
    haunted : bool,
    haunt_sink : Option<Sink>,
    haunt_stream : Option<OutputStream>,
    flicker_on: bool,
    last_flicker: Instant,
    corruption_level: u8,       // 0 = calm, 3 = full chaos
    last_corruption_tick: Instant,
    haunted_start: Option<Instant>,
    jump_scare_triggered: bool,


}

impl Spreadsheet {
    /// Creates a new `Spreadsheet` instance with the given number of rows and columns.
    ///
    /// This method initializes a spreadsheet with the specified dimensions, creating
    /// a grid of cells. It sets up the initial state for the spreadsheet, including the
    /// cursor position, mode, undo and redo stacks, and other related fields.
    ///
    /// # Arguments:
    /// - `rows`: The number of rows in the spreadsheet.
    /// - `cols`: The number of columns in the spreadsheet
    ///
    /// # Returns:
    /// A new `Spreadsheet` instance with the given number of rows and columns.
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
            haunted: false,
            haunt_sink: None,
            haunt_stream: None,
            flicker_on: false,
            last_flicker: Instant::now(),
            corruption_level: 0,
            last_corruption_tick: Instant::now(),
            haunted_start: None,
            jump_scare_triggered: false,
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

    /// Retrieves a reference to a cell at the given address.
    ///
    /// This method looks up a cell in the spreadsheet based on the provided address.
    ///
    /// # Arguments:
    /// - `addr`: A reference to the `CellAddress` of the cell to retrieve.
    ///
    /// # Returns:
    /// An `Option` containing a reference to the `Cell` if it exists, or `None` if the address is invalid.
    fn get_cell(&self, addr: &CellAddress) -> Option<&Cell> {
        self.data.get(&addr.to_string())
    }

     /// Retrieves a mutable reference to a cell at the given address.
    ///
    /// This method allows for modifying the cell at the specified address.
    ///
    /// # Arguments:
    /// - `addr`: A reference to the `CellAddress` of the cell to retrieve.
    ///
    /// # Returns:
    /// An `Option` containing a mutable reference to the `Cell` if it exists, or `None` if the address is invalid.
    fn get_cell_mut(&mut self, addr: &CellAddress) -> Option<&mut Cell> {
        self.data.get_mut(&addr.to_string())
    }

    /// Moves the cursor by the given number of columns and rows.
    ///
    /// This method updates the position of the cursor within the bounds of the spreadsheet.
    ///
    /// # Arguments:
    /// - `dx`: The number of columns to move the cursor. Positive values move to the right, negative values to the left.
    /// - `dy`: The number of rows to move the cursor. Positive values move down, negative values move up.
    ///
    /// # Notes:
    /// The cursor will not move outside the bounds of the spreadsheet (i.e., the number of columns and rows).

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

    /// Moves the cursor to the specified cell address.
    ///
    /// This method attempts to move the cursor to a given cell address, specified as a string (e.g., "A1").
    /// If the address is valid and within the bounds of the spreadsheet, the cursor will be moved to that cell.
    ///
    /// # Arguments:
    /// - `addr`: A string representing the cell address to jump to (e.g., "A1", "B2").
    ///
    /// # Returns:
    /// `true` if the cell address is valid and the cursor is successfully moved, otherwise `false`.
    fn jump_to_cell(&mut self, addr: &str) -> bool {
        if let Some(cell_addr) = CellAddress::from_str(addr) {
            if cell_addr.col < self.max_cols && cell_addr.row < self.max_rows {
                self.cursor = cell_addr;
                return true;
            }
        }
        false
    }

    /// Adds a dependency between two cells.
    ///
    /// This method records that one cell (the dependent) depends on the value of another cell (the dependency).
    /// It updates the `dependencies` and `dependents` mappings accordingly.
    ///
    /// # Arguments:
    /// - `dependent`: The address of the cell that depends on another cell.
    /// - `dependency`: The address of the cell that is being depended on.
    ///
    /// # Notes:
    /// This method will ensure that both the `dependencies` and `dependents` mappings are updated for both
    /// the dependent and the dependency cells.
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

    /// Removes all dependencies related to the given cell address.
    ///
    /// This method removes both the cell's dependencies and the cell from the list of dependents of each of its
    /// dependencies. It is useful when clearing dependencies when a cell's formula is changed or removed.
    ///
    /// # Arguments:
    /// - `cell_addr`: The address of the cell for which to remove dependencies.
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

    /// Updates the dependencies for a cell based on its formula.
    ///
    /// This method analyzes a cell's formula and updates its dependencies accordingly. The formula can refer to
    /// other cells directly (e.g., `A1`), ranges of cells (e.g., `SUM(A1:B2)`), or even functions with cell
    /// references (e.g., `=SUM(A1:B1)`).
    ///
    /// # Arguments:
    /// - `cell_addr`: The address of the cell whose dependencies need to be updated.
    /// - `formula`: The formula string that defines the dependencies.
    fn update_dependencies(&mut self, cell_addr: &str, formula: &str) {
        println!("DEBUG: Removing dependencies for cell {}", cell_addr);
        // First, remove any existing dependencies
        self.remove_dependencies(cell_addr);
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
    }
    /// Propagates changes through the spreadsheet based on cell dependencies.
    ///
    /// This method updates all the cells that depend on a given cell. If a cell's value changes, this method
    /// ensures that all dependent cells are recalculated. It also checks for circular dependencies and avoids
    /// infinite loops by tracking cells that are currently being updated.
    ///
    /// # Arguments:
    /// - `cell_addr`: A string representing the address of the cell whose changes need to be propagated.
    ///
    /// # Notes:
    /// - If a circular dependency is detected, an error message is shown, and the operation is undone.
    /// - This method processes each dependent cell recursively to ensure that the entire dependency chain is handled.
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
            let formula_opt = if let Some(cell) = self.data.get(&dependent) {
                cell.formula.clone()
            } else {
                None
            };
            if let Some(formula) = formula_opt {
                let formula_with_eq = format!("={}", formula);
                
                if let Some(addr) = CellAddress::from_str(&dependent) {
                    // Update the cell with its formula to recalculate
                    self.update_cell(&addr, &formula_with_eq, true);
                }
            }
        }
    }
    /// Updates a cell's value in the spreadsheet, recalculates it if necessary, and propagates changes
/// to dependent cells. This function supports both simple values and complex formulas (such as 
/// `SUM`, `MIN`, `MAX`, `sqrt`, and `log`). It also checks for circular dependencies and invalid 
/// formulas, ensuring that the integrity of the spreadsheet is maintained.
///
/// # Arguments
///
/// * `addr` - A reference to the `CellAddress` of the cell to be updated. This indicates which 
///   cell in the spreadsheet should be modified.
/// * `value` - A string representing the new value or formula for the cell. If the value starts 
///   with `=`, it is considered a formula; otherwise, it's treated as a constant value.
/// * `multi` - A boolean flag indicating whether this update is part of a multi-cell operation. 
///   If `multi` is `false`, the function will push the current state to the undo stack to allow 
///   for future undo operations. If `multi` is `true`, undo history will not be updated.
///
/// # Returns
///
/// Returns `true` if the cell was updated successfully, and `false` if an error occurred (e.g., 
/// invalid formula, circular dependency, or locked cell).
///
/// # Error Handling
///
/// This function performs several checks and sets the `status_message` with an appropriate error 
/// message if any of the following conditions are met:
/// 
/// - The cell doesn't exist (`ERROR: CELL {addr} NOT FOUND`)
/// - The cell is locked (`ERROR: CELL {addr} LOCKED`)
/// - A circular dependency is detected (`ERROR: CIRCULAR DEPENDENCY DETECTED EARLY WITH {addr}`)
/// - An invalid formula is provided, such as an incorrectly formatted range (`ERROR: INVALID RANGE {range}`)
/// - An invalid arithmetic expression (`ERROR: INVALID ARITHMETIC EXPRESSION {expression}`)
/// - An invalid function argument (`ERROR: INVALID ARGUMENT {function}`)
/// - A general invalid formula error (`ERROR: INVALID FORMULA {value}`)
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
        if let Some(_old_cell) = self.get_cell(addr).cloned() {

            let is_valid_formula: bool;
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
                            if !(start_exists && end_exists) {
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
                    }
                    else if cell_ref.contains('+') || cell_ref.contains('-') || cell_ref.contains('*') {
                        // Arithmetic expression like =(A1+B1)
                        let re = regex::Regex::new(r"([+\-*])").unwrap();
                        let parts: Vec<&str> = re.split(cell_ref).collect();
                        
                        // Check if all parts are valid (either cell references or numbers)
                        let all_valid = parts.iter().all(|part| {
                            let trimmed = part.trim();
                            if trimmed.is_empty() {
                                return false;
                            }
                            
                            // Check if it's a valid cell reference
                            if let Some(addr) = CellAddress::from_str(trimmed) {
                                self.get_cell(&addr).is_some()
                            } else {
                                // Check if it's a valid number
                                trimmed.parse::<f64>().is_ok()
                            }
                        });
                        
                        if !all_valid {
                            self.status_message = format!("ERROR: INVALID ARITHMETIC EXPRESSION {}", cell_ref);
                            false
                        } else {
                            true
                        }
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
                if !multi{
                    println!("DEBUG: Pushing undo for cell {}", addr.to_string());
                    self.push_undo_sheet();
                    self.redo_stack.clear(); 
                }
                // self.push_undo_sheet();
                // self.redo_stack.clear(); 

                self.update_dependencies(&addr.to_string(), value);

                if let Some(cell) = self.get_cell_mut(addr) {
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
                if !multi{
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
                    let inside_brackets = &formula[1..formula.len() - 1];
                    
                    if let Some(addr) = CellAddress::from_str(inside_brackets) {
                        // Simple cell reference like =(A1)
                        println!("DEBUG: Found simple cell reference in formula");
                        if let Some(cell) = self.get_cell(&addr) {
                            if let Ok(value) = cell.display_value.parse::<f64>() {
                                value
                            } else {
                                0.0
                            }
                        } else {
                            0.0
                        }
                    } else if inside_brackets.contains('+') || inside_brackets.contains('-') || inside_brackets.contains('*') {
                        // Arithmetic expression like =(A1+B1) or =(A1+1)
                        println!("DEBUG: Found arithmetic expression in formula: {}", inside_brackets);
                        
                        // Find the operator and its position
                        let mut operator = '+';  // Default
                        let mut operator_pos = 0;
                        
                        for (i, c) in inside_brackets.chars().enumerate() {
                            if c == '+' || c == '-' || c == '*' {
                                operator = c;
                                operator_pos = i;
                                break;
                            }
                        }
                        
                        let left_part = &inside_brackets[0..operator_pos].trim();
                        let right_part = &inside_brackets[operator_pos+1..].trim();
                        
                        // Evaluate left operand
                        let left_value = if let Some(addr) = CellAddress::from_str(left_part) {
                            if let Some(cell) = self.get_cell(&addr) {
                                cell.display_value.parse::<f64>().unwrap_or(0.0)
                            } else {
                                0.0
                            }
                        } else {
                            left_part.parse::<f64>().unwrap_or(0.0)
                        };
                        
                        // Evaluate right operand
                        let right_value = if let Some(addr) = CellAddress::from_str(right_part) {
                            if let Some(cell) = self.get_cell(&addr) {
                                cell.display_value.parse::<f64>().unwrap_or(0.0)
                            } else {
                                0.0
                            }
                        } else {
                            right_part.parse::<f64>().unwrap_or(0.0)
                        };
                        
                        // Perform the operation
                        match operator {
                            '+' => left_value + right_value,
                            '-' => left_value - right_value,
                            '*' => left_value * right_value,
                            _ => 0.0  // Should not reach here due to validation
                        }
                    } else {
                        println!("DEBUG: Invalid content in brackets: {}", inside_brackets);
                        0.0
                    }
                }
                else {
                    0.0
                };
                // Update the cell's display value with the computed result
                if let Some(cell) = self.get_cell_mut(addr) {
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

    // Pushes a single undo action to the undo stack for a specific cell update. This action stores
// the previous state of the cell so that it can be reverted during an undo operation.
//
// The undo stack is capped at 3 actions, and older actions are discarded when this limit is exceeded.
//
// # Arguments
//
// * `addr` - The `CellAddress` of the cell that was updated.
// * `old_cell` - A `Cell` representing the state of the cell before the update.
//
// # Notes
//
// The undo stack is maintained in a way that only a limited number of undo actions are stored
// at any given time. If the stack reaches its limit, the oldest action is discarded to make room
// for new actions.
    // fn push_undo(&mut self, addr: CellAddress, old_cell: Cell) {
    //     // Maintain max 3 undo steps
    //     if self.undo_stack.len() >= 3 {
    //         self.undo_stack.pop_front();
    //     }
    //     self.undo_stack.push_back(UndoAction {
    //         cell_address: addr,
    //         old_cell
    //     });
    // }

    /// Pushes the entire sheet's state to the undo stack. This operation adds all current cells in
/// the sheet to the undo stack so that the entire sheet can be reverted in a single undo operation.
///
/// The undo stack is capped at 3 actions, and older actions are discarded when this limit is exceeded.
/// If the undo stack already contains 3 actions, it is cleared before adding a new action.
///
/// # Example
///
/// # Notes
///
/// This operation clears the undo stack when adding the first action if the cell at address `A1`
/// is present in the data and the undo stack already has 3 actions.
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
    /// Undoes the last action applied to the sheet. If the undo stack is empty, a message is set
/// indicating that there is nothing to undo.
///
/// The state of the sheet is reverted to the state it was in before the last action. The undone
/// actions are then moved to the redo stack, allowing them to be reapplied later using the redo function.
///
/// # Returns
///
/// Returns `true` if the undo operation was successfully applied, or `false` if there was nothing to undo.
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
    /// Redoes the last undone action. If the redo stack is empty, a message is set indicating that
/// there is nothing to redo.
///
/// The state of the sheet is restored to the state it was in before the undo operation. The redone
/// actions are then moved back to the undo stack, allowing them to be undone again if needed.
///
/// # Returns
///
/// Returns `true` if the redo operation was successfully applied, or `false` if there was nothing to redo.
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

    /// Locks a specific cell, preventing its value from being modified until it is unlocked.
/// If no address is provided, the currently selected cell (cursor) will be locked.
///
/// # Arguments
///
/// * `addr` - An optional string slice representing the cell's address to be locked. If not provided,
///   the currently selected cell is locked.
///
/// # Returns
///
/// Returns `true` if the cell was successfully locked, or `false` if the cell could not be locked 
/// (e.g., invalid address).
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
/// Unlocks a specific cell, allowing its value to be modified. If no address is provided, 
/// the currently selected cell (cursor) will be unlocked.
///
/// # Arguments
///
/// * `addr` - An optional string slice representing the cell's address to be unlocked. If not provided,
///   the currently selected cell is unlocked.
///
/// # Returns
///
/// Returns `true` if the cell was successfully unlocked, or `false` if the cell could not be unlocked 
/// (e.g., invalid address).
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
/// Sets the alignment of a specific cell. The alignment can be set to left, right, or center.
/// If no address is provided, the currently selected cell (cursor) will be modified.
///
/// # Arguments
///
/// * `addr` - An optional string slice representing the cell's address. If not provided,
///   the currently selected cell is used.
/// * `align` - A string that specifies the alignment. Possible values are:
///   - `"l"` for left alignment
///   - `"r"` for right alignment
///   - `"c"` for center alignment
///
/// # Returns
///
/// Returns `true` if the alignment was successfully changed, or `false` if the address is invalid,
/// the cell is locked, or the alignment value is invalid.
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
/// Sets the height and width for a specific cell. If no address is provided, the currently selected 
/// cell (cursor) will be modified. The height and width can be adjusted independently.
///
/// # Arguments
///
/// * `addr` - An optional string slice representing the cell's address. If not provided,
///   the currently selected cell is used.
/// * `height` - An optional `usize` representing the height of the cell. If not provided, the height
///   will not be changed.
/// * `width` - An optional `usize` representing the width of the cell. If not provided, the width
///   will not be changed.
///
/// # Returns
///
/// Returns `true` if the dimension was successfully changed, or `false` if the address is invalid,
/// the cell is locked, or invalid dimensions were provided.
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
/// Searches for a query string within all cells in the spreadsheet. If any cells contain the query,
/// their addresses will be stored as matches.
///
/// # Arguments
///
/// * `query` - The string to search for in the cell values.
///
/// # Returns
///
/// Returns `true` if one or more matches are found, and sets the cursor to the first match. 
/// Returns `false` if no matches are found.
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
/// Navigates to the next matching cell in the find results. The cursor will be updated to the next
/// match in the list of search results.
///
/// # Returns
///
/// Returns `true` if a match is found and the cursor is updated. Returns `false` if no matches have been found.
    fn find_next(&mut self) -> bool {
        if self.find_matches.is_empty() {
            return false;
        }
        
        self.current_find_match = (self.current_find_match + 1) % self.find_matches.len();
        self.cursor = self.find_matches[self.current_find_match].clone();
        true
    }
/// Navigates to the previous matching cell in the find results. The cursor will be updated to the previous
/// match in the list of search results.
///
/// # Returns
///
/// Returns `true` if a match is found and the cursor is updated. Returns `false` if no matches have been found.
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

    /// Parses a range string in the format "A1:B5" into two `CellAddress` objects representing
/// the starting and ending cell addresses. If the format is invalid, returns `None`.
///
/// # Arguments
///
/// * `range_str` - A string representing the range to parse (e.g., "A1:B5").
///
/// # Returns
///
/// Returns an `Option` containing a tuple of `CellAddress` objects for the start and end cells if valid,
/// or `None` if the format is invalid or the cell addresses cannot be parsed.
    fn parse_range(&self, range_str: &str) -> Option<(CellAddress, CellAddress)> {
        let parts: Vec<&str> = range_str.split(':').collect();
        if parts.len() != 2 {
            return None;
        }
        
        let start = CellAddress::from_str(parts[0])?;
        let end = CellAddress::from_str(parts[1])?;
        
        Some((start, end))
    }
/// Inserts a specified value into a range of cells. The range is parsed from the `range_str`
/// argument (e.g., "A1:B3"), and the value is inserted into all cells within that range. 
/// The undo stack is updated before any changes are made.
///
/// # Arguments
///
/// * `range_str` - A string representing the range to insert the value into (e.g., "A1:B3").
/// * `value` - The value to insert into the specified range of cells.
///
/// # Returns
///
/// Returns `true` if the value was successfully inserted into the specified range, or `false` if:
/// - The range is invalid.
/// - Any of the cells in the range are locked (the update will skip locked cells).
/// - An error occurs while processing the range.
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
/// Saves the current spreadsheet data as a JSON file to the specified path.
///
/// # Arguments
///
/// * `path` - The path where the JSON file should be saved.
///
/// # Returns
///
/// Returns `io::Result<()>`, which will be `Ok` if the file is written successfully, or an error if
/// there is an issue with creating or writing to the file.
    fn save_json(&self, path: &Path) -> io::Result<()> {
        let file = File::create(path)?;
        let writer = BufWriter::new(file);
        serde_json::to_writer_pretty(writer, &self.data)?;
        Ok(())
    }
/// Loads spreadsheet data from a JSON file at the specified path.
///
/// # Arguments
///
/// * `path` - The path to the JSON file containing the spreadsheet data.
///
/// # Returns
///
/// Returns `io::Result<()>`, which will be `Ok` if the file is read and the data is successfully loaded,
/// or an error if the file cannot be opened or the data cannot be parsed.
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
/// Sorts the rows within a specified range of cells based on the values in a given column. The rows
/// can be sorted in either ascending or descending order.
///
/// # Arguments
///
/// * `range_str` - A string representing the range to sort (e.g., "A1:B5").
/// * `ascending` - A boolean flag indicating the sort order. `true` for ascending, `false` for descending.
///
/// # Returns
///
/// Returns `true` if the sorting operation was successful, or `false` if:
/// - The range is invalid.
/// - An error occurs during the sorting process.
///
/// # Notes
///
/// The function performs the following steps:
/// 1. Extracts the range of cells to be sorted from the provided string.
/// 2. Sorts the rows based on the values in the specified column, comparing first by numeric value (if possible),
///    and then by string value.
/// 3. Applies the sorted rows back to the sheet.
/// 4. The undo stack is updated before sorting, and the redo stack is cleared.
///
/// If a cell is locked, it will not be modified during the sorting operation.
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
/// Formats the value of a cell for display, taking into account its width and alignment.
///
/// # Arguments
///
/// * `addr` - A reference to the `CellAddress` of the cell whose value is to be formatted.
///
/// # Returns
///
/// Returns a `String` containing the formatted cell value, which may be truncated to fit the cell's width
/// and padded according to the specified alignment (left, right, or center).
/// # Notes
///
/// This function formats the value of the cell to fit within the defined width:
/// - If the cell's value exceeds its width, it will be truncated with an ellipsis (`..`) if there's enough space.
/// - The cell's value will be padded with spaces based on its alignment (left, right, or center).
///
/// If the width is too small to display any part of the value, the cell will display a series of periods (`"."`).
    fn format_cell_value(&self, addr: &CellAddress) -> String {
        let cell = self.get_cell(addr).clone().unwrap(); 
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
/// Exports the spreadsheet data to a PDF file with formatted content including rows, columns, and cell values.
///
/// The export includes the following features:
/// - Data from the spreadsheet is formatted in a table-like structure with row numbers and column headers.
/// - Content is split into multiple pages if there are more rows than can fit on one page.
/// - Page numbers are included in the footer (e.g., "Page 1 of 3").
///
/// # Arguments
///
/// * `filename` - The name of the output PDF file. This is where the PDF will be saved.
///
/// # Returns
///
/// Returns a `Result<(), io::Error>`. On success, it returns `Ok(())`. On failure, it returns an `Err`
/// with the error details.
///
/// # Notes
///
/// This function does the following:
/// 1. Creates a new PDF document with A4 page dimensions.
/// 2. Iterates through the spreadsheet data and splits it across pages if needed.
/// 3. Draws the column headers and row numbers on each page.
/// 4. Writes the cell values within the table format, considering cell width and row height.
/// 5. Adds page numbers to the bottom of each page (e.g., "Page X of Y").
/// 6. Saves the PDF document to the provided file path.
///
/// The resulting PDF will have the following layout:
/// - Each page shows a part of the table with row numbers on the left, followed by columns A to J.
/// - The table content will be truncated if the width of the columns exceeds the page width.
/// - The rows will be adjusted to fit within the available content height on each page.
    fn export_to_pdf(&self, filename: &str) -> Result<()> {
        // Create a new PDF document
        let ( doc, page1, layer1) = PdfDocument::new("Spreadsheet Export", Mm(210.0), Mm(297.0), "Layer 1");
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
/// Processes and executes a command entered by the user.
///
/// This function interprets a variety of user commands, changing the state of the spreadsheet 
/// or performing specific actions based on the command entered. It handles commands related to 
/// cell manipulation, navigation, file operations, and special features like "haunting".
///
/// # Command Syntax
/// - Commands can be entered in the form of single-letter abbreviations or with parameters, 
///   e.g., `i cell_name`, `find search_term`, `sort range ascending_flag`, etc.
/// - Commands that are not recognized will display an "INVALID COMMAND" status message.
///
/// # Command List
/// - `"q"`: Quit the application.
/// - `"i [cell]"`: Enter insert mode at the specified cell (or current cell if no cell specified).
/// - `"j [cell]"`: Jump to the specified cell.
/// - `"undo"`: Undo the last operation.
/// - `"redo"`: Redo the last undone operation.
/// - `"find [search_term]"`: Enter find mode with the specified search term.
/// - `"mi [start] [end]"`: Multi-insert command for a range of values.
/// - `"lock [cell]"`: Lock the specified cell, or lock the current cell if no cell is specified.
/// - `"unlock [cell]"`: Unlock the specified cell, or unlock the current cell if no cell is specified.
/// - `"align [alignment]"`: Set alignment for the current cell or a specified cell.
/// - `"dim [cell] (height,width)"`: Set dimensions (height and width) for a cell.
/// - `"sort [range] [ascending_flag]"`: Sort a range of cells in ascending or descending order.
/// - `"saveas_<format> [filename]"`: Save the spreadsheet as the specified format (e.g., JSON or PDF).
/// - `"load [filename]"`: Load a spreadsheet from a file.
/// - `"hh"`: Go to the leftmost cell in the current row.
/// - `"ll"`: Go to the rightmost cell in the current row.
/// - `"jj"`: Go to the bottommost cell in the current column.
/// - `"kk"`: Go to the topmost cell in the current column.
/// - `"haunt"`: Enable haunting mode, play a sound, and display a haunting message.
/// - `"dehaunt"`: Disable haunting mode and stop the sound if it's playing.
///
/// # Arguments
///
/// This function takes no arguments but relies on the `command_buffer` property of the struct to
/// capture the user input.
///
/// # Returns
///
/// Returns a boolean value, always `true`, indicating that the process will continue running 
/// unless the user enters the "q" command (which causes the function to return `false`).
///
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
        }  else if cmd == "haunt" {
            self.haunted = true;
            self.haunted_start = Some(Instant::now());
            self.jump_scare_triggered = false;
        
            // WSL-friendly sound playback
            let windows_path = r#"C:\Users\hp\OneDrive - IIT Delhi\Desktop\Academics\prisha_rust_lab\creaking_door.wav"#; 
            play_sound(windows_path);
        
            self.status_message = "👻 You are being haunted...".to_string();
        } else if cmd == "dehaunt" {
            self.haunted = false;
            self.haunted_start = None;
            self.jump_scare_triggered = false;
        
            if let Some(sink) = &self.haunt_sink {
                sink.stop(); // stop playback
            }
        
            self.haunt_sink = None;
            self.haunt_stream = None;
            self.status_message = "🧹 Haunting ended.".to_string();
        } else {
            self.status_message = "INVALID COMMAND".to_string();
        }
        
        true // Continue running
    }
/// Handles key events based on the current mode of the application.
///
/// This function processes the key presses based on the current mode of the application 
/// (Normal, Insert, Command, or Find mode). It handles cursor movements, inserting values, 
/// running commands, and more.
///
/// # Mode Behavior
/// - **Normal Mode**: 
///     - `h`, `j`, `k`, `l` to move the cursor left, down, up, and right respectively.
///     - `w`, `a`, `s`, `d` to scroll the view.
///     - `:` to switch to Command Mode.
///     - `q` to quit the application.
/// - **Insert Mode**: 
///     - `Esc` to switch back to Normal Mode.
///     - `Enter` to apply the changes to the cell and return to Normal Mode.
///     - `Backspace` to remove the last character from the command buffer.
///     - Any character is inserted into the command buffer.
/// - **Command Mode**: 
///     - `Esc` to return to Normal Mode.
///     - `Enter` to execute the command from the buffer and return to Normal Mode.
///     - `Backspace` to remove the last character from the command buffer.
///     - Any character is added to the command buffer.
/// - **Find Mode**: 
///     - `Esc` to return to Normal Mode and clear the find matches.
///     - `n` to find the next match.
///     - `p` to find the previous match.
///
/// # Arguments
/// 
/// * `key` - The key that was pressed (of type `KeyCode`), which is processed based on the current mode.
///
/// # Returns
/// 
/// Returns a boolean value:
/// - `true` to continue running the application.
/// - `false` if the user pressed `q` in Normal Mode (to quit the application).
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
                        if START_COL + 20 <= C - 1 {
                            START_COL += 10;
                        } else {
                            START_COL =  C.saturating_sub(10);
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
                        if START_ROW + 20 <= R - 1 {
                            START_ROW += 10;
                        } else {
                            START_ROW = R.saturating_sub(10);
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
    /// Draws the spreadsheet grid and related UI elements to the terminal.
///
/// This function is responsible for rendering the spreadsheet's grid, including:
/// - Row and column headers
/// - The cells' contents, with appropriate formatting and spacing
/// - The cursor position (highlighted)
/// - A status bar that shows information about the current cell
/// - A status message and command buffer, if available
///
/// The screen is cleared at the beginning, and the entire grid is redrawn, ensuring
/// that any changes (like cursor movement, updated cell content, or mode transitions) 
/// are reflected in the UI. This function works with a terminal-based UI and uses 
/// `terminal` and `cursor` functionalities to move the cursor and clear the screen.
///
/// # Arguments
/// 
/// * `stdout` - The output stream for writing terminal content, typically the terminal's standard output.
/// 
/// # Returns
/// 
/// Returns an `io::Result<()>`:
/// - `Ok(())` if the drawing was successful.
/// - `Err(e)` if an I/O error occurred during the process.


fn draw(&mut self, stdout: &mut io::Stdout) -> io::Result<()> {
    use rand::Rng;

    // Flicker toggle every 300ms
    if self.haunted && self.last_flicker.elapsed() > Duration::from_millis(300) {
        self.flicker_on = !self.flicker_on;
        self.last_flicker = Instant::now();
    }
    // Corruption increases every 5 seconds while haunted
    if self.haunted && self.last_corruption_tick.elapsed() > Duration::from_secs(7) {
        self.corruption_level = self.corruption_level.saturating_add(1).min(3);
        self.last_corruption_tick = Instant::now();
    }


    // Clear screen
    stdout.execute(terminal::Clear(ClearType::All))?;
    stdout.execute(MoveTo(0, 0))?;
    
    let row_label_width = 5;
    let cell_padding = 1;
    let default_cell_width = 5;
    let mut col_widths = vec![default_cell_width; 10];

    for col in unsafe { START_COL..(START_COL + 10) } {
        let col_idx = (col - unsafe { START_COL }) as usize;
        let col_letter = CellAddress::col_to_letters(col);
        col_widths[col_idx] = col_widths[col_idx].max(col_letter.len());
        for row in unsafe { START_ROW..(START_ROW + 10).min(R) } {
            let addr = CellAddress::new(col, row);
            if let Some(cell) = self.get_cell(&addr) {
                col_widths[col_idx] = col_widths[col_idx].max(cell.width);
            }
        }
        col_widths[col_idx] = col_widths[col_idx].max(3);
    }

    stdout.execute(SetForegroundColor(Color::Cyan))?;
    write!(stdout, "{:<width$}", "", width = row_label_width + 1)?;

    for col in unsafe { START_COL..(START_COL + 10).min(C) } {
        let col_idx = (col - unsafe { START_COL }) as usize;
        let col_letter = CellAddress::col_to_letters(col);
        let total_cell_width = col_widths[col_idx] + cell_padding;
        write!(stdout, "{:^width$}", col_letter, width = total_cell_width)?;
    }

    write!(stdout, "\r\n")?;

    if self.haunted && rand::random::<u8>() % 100 == 0 {
        stdout.execute(SetForegroundColor(Color::Red))?;
        write!(stdout, "{}", "👻")?;
        stdout.execute(SetForegroundColor(Color::Reset))?;
    }

    let mut rng = rand::thread_rng();

    for row in unsafe { START_ROW..(START_ROW + 10).min(R) } {
        stdout.execute(SetForegroundColor(Color::Cyan))?;
        write!(stdout, "{:>width$}", row + 1, width = row_label_width)?;
        stdout.execute(SetForegroundColor(Color::Reset))?;

        for col in unsafe { START_COL..(START_COL + 10).min(C) } {
            let col_idx = (col - unsafe { START_COL }) as usize;
            let addr = CellAddress::new(col, row);
            let is_cursor_cell = col == self.cursor.col && row == self.cursor.row;

            // Haunted flicker logic
            let mut flicker_effect = None;

            if self.haunted && self.flicker_on {
                let chance: f32 = rng.r#gen();

                match self.corruption_level {
                    0 => {
                        if chance < 0.05 {
                            flicker_effect = Some("👻");
                        }
                    }
                    1 => {
                        if chance < 0.05 {
                            flicker_effect = Some("👻");
                        } else if chance < 0.10 {
                            flicker_effect = Some("~");
                        }
                    }
                    2 => {
                        if chance < 0.05 {
                            flicker_effect = Some("👻");
                        } else if chance < 0.10 {
                            flicker_effect = Some(["~", "#", "X", "%", "!!"].choose(&mut rng).unwrap());
                        } else if chance < 0.12 {
                            flicker_effect = Some("💥");
                        }
                    }
                    3 => {
                        if chance < 0.05 {
                            flicker_effect = Some("👻");
                        } else if chance < 0.10 {
                            flicker_effect = Some(["~", "#", "X", "%", "!!", "???"].choose(&mut rng).unwrap());
                        } else if chance < 0.15 {
                            flicker_effect = Some("💥");
                        }
                    }
                    _ => {}
                }
            }

            if self.haunted && self.corruption_level >= 2 && rng.r#gen::<f32>() < 0.02 {
                let whispers = [
                    "get out",
                    "it sees you",
                    "run",
                    "don't trust it",
                    "they're watching",
                    "help me",
                    "leave now",
                ];
                self.status_message = whispers.choose(&mut rng).unwrap().to_string();
            }
            



            // Handle flicker color
            // if flicker_dim {
            //     stdout.execute(SetForegroundColor(Color::DarkGrey))?;
            // }

            // Cursor highlight
            if is_cursor_cell {
                stdout.execute(SetForegroundColor(Color::Black))?;
                stdout.execute(style::SetBackgroundColor(Color::White))?;
            }

            let _cell_content = if let Some(cell) = self.get_cell(&addr) {
                cell.display_value.clone()
            } else {
                "0".to_string()
            };

            let _available_width = col_widths[col_idx];
            // if cell_content.len() > available_width {
            //     cell_content = format!("{}..", &cell_content[0..available_width.saturating_sub(2)]);
            // }

            // Draw or skip content based on flicker
            if let Some(effect) = flicker_effect {
                // Extra chaos: highlight 💥 in red
                if effect == "💥" {
                    stdout.execute(SetForegroundColor(Color::Red))?;
                    stdout.execute(style::SetBackgroundColor(Color::Black))?;
                }
                write!(stdout, " {:^width$}", effect, width = col_widths[col_idx])?;
                stdout.execute(SetForegroundColor(Color::Reset))?;
                stdout.execute(style::SetBackgroundColor(Color::Reset))?;
            } else {
                write!(stdout, " {:^width$}", self.format_cell_value(&addr), width = col_widths[col_idx])?;
            }
            
            

            // Reset styles
            if is_cursor_cell {
                stdout.execute(SetForegroundColor(Color::Reset))?;
                stdout.execute(style::SetBackgroundColor(Color::Reset))?;
            }

            // if flicker_dim {
            //     stdout.execute(SetForegroundColor(Color::Reset))?;
            // }
        }

        write!(stdout, "\r\n")?;
    }

    writeln!(stdout)?;

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

    let (cols, rows) = terminal::size()?;
    let status_message = &self.status_message;
    if !status_message.is_empty() {
        stdout.execute(MoveTo(cols.saturating_sub(status_message.len() as u16), rows.saturating_sub(1)))?;
        write!(stdout, "{}", status_message)?;
    }

    if !self.command_buffer.is_empty() {
        let command_buffer = &self.command_buffer;
        stdout.execute(MoveTo(0, rows.saturating_sub(2)))?;
        write!(stdout, "{}", command_buffer)?;
    }

    stdout.flush()?;

    Ok(())
}
}

/// Main function to initialize and run the extended spreadsheet application.
///
/// This function sets up the terminal in raw mode and creates a spreadsheet with a configurable
/// size (determined by command-line arguments). It runs a main event loop where the current state
/// of the spreadsheet is drawn and user input is handled to manipulate the spreadsheet. The loop
/// continues until the user decides to quit. Once the program ends, the terminal is cleaned up, 
/// and the cursor is restored.
///
/// # Command-Line Arguments
/// 
/// The program expects two command-line arguments:
/// - `<rows>`: The number of rows in the spreadsheet. Defaults to `10` if not provided.
/// - `<cols>`: The number of columns in the spreadsheet. Defaults to `10` if not provided.
/// 
/// If the number of arguments provided is incorrect, the program will display an error message and
/// default to a 10x10 grid.
///
/// # Behavior
/// - The terminal is cleared, raw mode is enabled, and the cursor is hidden to allow custom rendering.
/// - The event loop waits for key events to handle user input (e.g., navigating the spreadsheet or editing cells).
/// - The loop continues until the user exits (via the `handle_key_event` method returning `false`).
/// - Upon exit, the terminal is restored, the cursor is shown again, and the screen is cleared.
///
/// # Terminal Settings
/// - Raw mode is enabled with `terminal::enable_raw_mode()`, which allows direct control over input and output.
/// - The cursor is hidden initially and shown again upon exit to maintain the custom UI.
pub fn main() -> Result<()> {
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
    stdout.execute(Hide)?; // Hide cursor for custom rendering

    // Create spreadsheet (10x10 grid)
    let mut sheet = Spreadsheet::new(rows, cols);

    // Main event loop
    loop {
        // Draw the current state
        if sheet.haunted {
            if let Some(start_time) = sheet.haunted_start {
                if !sheet.jump_scare_triggered && start_time.elapsed() > Duration::from_secs(15) {
                    trigger_jump_scare();
                    sheet.jump_scare_triggered = true;
                    std::thread::sleep(std::time::Duration::from_secs(2));
                    sheet.jump_scare_triggered = true;
                }
            }
        }
        
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
    stdout.execute(Show)?; // Show cursor again
    stdout.execute(terminal::Clear(ClearType::All))?;
    stdout.execute(MoveTo(0, 0))?;

    Ok(())
}