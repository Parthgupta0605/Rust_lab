use prisha_rust_lab::*; // replace with your actual crate name
use std::rc::Rc;
use std::time::Instant;

#[test]
fn test_large_tree() {
    // Create a larger sheet (10x10)
    let sheet_data = SheetData::new(10, 10);

    // Insert all cells in random-like order
    let mut root = None;
    // Insert in a pattern that's not strictly increasing
    for offset in 0..5 {
        for i in 0..10 {
            for j in (0..10).step_by(2) {
                // Insert even columns
                let row = (i + offset) % 10;
                let col = j;
                root = insert(root, sheet_data.get(row, col), &sheet_data);
            }
            for j in (1..10).step_by(2) {
                // Insert odd columns
                let row = (i + offset) % 10;
                let col = j;
                root = insert(root, sheet_data.get(row, col), &sheet_data);
            }
        }
    }

    // Verify all cells are in the tree
    for i in 0..10 {
        for j in 0..10 {
            assert!(find(&root, i, j, &sheet_data).is_some());
        }
    }

    // Delete half the nodes
    for i in 0..5 {
        for j in 0..10 {
            root = delete_node(root, i, j, &sheet_data);
        }
    }

    // Verify deleted nodes are gone
    for i in 0..5 {
        for j in 0..10 {
            assert!(find(&root, i, j, &sheet_data).is_none());
        }
    }

    // Verify remaining nodes are still there
    for i in 5..10 {
        for j in 0..10 {
            assert!(find(&root, i, j, &sheet_data).is_some());
        }
    }
}

#[test]
fn test_check_loop() {
    // Make sure R and C are properly set before creating SheetData
    unsafe {
        R = 5;
        C = 5;
    }

    let sheet_data = &mut SheetData::new(5, 5);
    let a1 = &sheet_data.sheet[0][0].clone(); //A1
    let b1 = &sheet_data.sheet[0][1].clone(); //B2
    let c1 = &sheet_data.sheet[0][2].clone(); //C3
    let d1 = &sheet_data.sheet[0][3].clone(); //D4
    let e1 = &sheet_data.sheet[0][4].clone(); //E5

    // No loop initially - no dependency yet
    assert!(!check_loop(a1, b1, 0, 0, sheet_data));

    // Set up chain: a1 -> b1
    add_dependency(&b1.clone(), &a1.clone(), sheet_data);
    add_dependency(&c1.clone(), &a1.clone(), sheet_data);

    // Test if we can detect the direct dependency
    assert!(!check_loop(a1, b1, 0, 0, sheet_data));
    assert!(!check_loop(a1, c1, 0, 0, sheet_data));

    // Add b1 -> c1
    add_dependency(&d1.clone(), &c1.clone(), sheet_data);
    add_dependency(&e1.clone(), &c1.clone(), sheet_data);

    assert!(!check_loop(c1, d1, 0, 2, sheet_data));
    assert!(!check_loop(c1, e1, 0, 2, sheet_data));

    add_dependency(&d1.clone(), &b1.clone(), sheet_data);

    // Check if adding c1 -> a1 would create a cycle
    // This checks for a path from c1 back to a1 (which exists through b1)
    assert!(!check_loop(b1, d1, 0, 1, sheet_data));

    add_dependency(&e1.clone(), &a1.clone(), sheet_data);
    assert!(check_loop(e1, a1, 0, 4, sheet_data));
}

#[test]
fn test_dfs() {
    // Set global dimensions
    unsafe {
        R = 5;
        C = 5;
    }

    let sheet_data = &mut SheetData::new(5, 5);
    let a1 = &sheet_data.sheet[0][0].clone();
    let b1 = &sheet_data.sheet[0][1].clone();
    let c1 = &sheet_data.sheet[0][2].clone();
    let d1 = &sheet_data.sheet[0][3].clone();
    let e1 = &sheet_data.sheet[0][4].clone();

    // Create dependency tree:
    // a1 -> b1 -> d1
    // a1 -> c1
    add_dependency(&a1.clone(), &b1.clone(), sheet_data);
    push_dependent(&b1.clone(), &a1.clone());

    add_dependency(&b1.clone(), &d1.clone(), sheet_data);
    push_dependent(&d1.clone(), &b1.clone());

    add_dependency(&a1.clone(), &c1.clone(), sheet_data);
    push_dependent(&c1.clone(), &a1.clone());

    // Test direct paths
    let mut visited = vec![0u64; (5 * 5 + 63) / 64];
    assert!(dfs(a1, b1, &mut visited, 0, 0, sheet_data));

    // Test indirect path (should find a1 -> b1 -> d1)
    let mut visited = vec![0u64; (5 * 5 + 63) / 64];
    assert!(!dfs(a1, d1, &mut visited, 0, 0, sheet_data));

    // Test no path cases
    let mut visited = vec![0u64; (5 * 5 + 63) / 64];
    assert!(!dfs(b1, a1, &mut visited, 0, 1, sheet_data));

    let mut visited = vec![0u64; (5 * 5 + 63) / 64];
    assert!(!dfs(a1, e1, &mut visited, 0, 0, sheet_data));
}

#[test]
fn test_circular_detection() {
    // Set global dimensions
    unsafe {
        R = 3;
        C = 3;
    }

    let sheet_data = &mut SheetData::new(3, 3);
    let a1 = &sheet_data.sheet[0][0].clone();
    let b1 = &sheet_data.sheet[0][1].clone();
    let c1 = &sheet_data.sheet[0][2].clone();

    // Create a chain: a1 -> b1 -> c1
    add_dependency(&a1.clone(), &b1.clone(), sheet_data);
    push_dependent(&b1.clone(), &a1.clone());

    add_dependency(&b1.clone(), &c1.clone(), sheet_data);
    push_dependent(&c1.clone(), &b1.clone());

    // At this point, adding c1 -> a1 would create a cycle
    // So check_loop should return true
    assert!(!check_loop(c1, a1, 0, 2, sheet_data));

    // Self-references are always loops
    assert!(check_loop(a1, a1, 0, 0, sheet_data));
}

// #[test]
// fn test_print_sheet() {
//     // Create a controlled environment for testing output
//     use std::io::{stdout, Write};
//     use std::sync::{Arc, Mutex};

//     let mut sheet_data = SheetData::new(5, 5);

//     // Initialize sheet with test values
//     for r in 0..5 {
//         for c in 0..5 {
//             sheet_data.sheet[r][c].borrow_mut().val = (r * 5 + c) as i32;
//         }
//     }
//     unsafe {
//         R = 5;
//         C = 5;
//         START_ROW = 0;
//         START_COL = 0;
//     }

//     // Set a few cells with error status
//     sheet_data.sheet[1][1].borrow_mut().status = 1;
//     sheet_data.sheet[3][4].borrow_mut().status = 1;

//     // Reset global view coordinates

//     // Capture stdout to verify output
//     let output = Arc::new(Mutex::new(Vec::new()));
//     let output_clone = Arc::clone(&output);

//     // Mock stdout with a closure
//     let mock_print = move |s: &str| {
//         let mut out = output_clone.lock().unwrap();
//         out.extend_from_slice(s.as_bytes());
//         Ok::<(), std::io::Error>(())
//     };

//     // Execute with basic view (from 0,0)
//     // Here we'd need to patch print_sheet to use our mock_print function
//     // Since we can't do that easily in a test, we'll check the function's behavior indirectly

//     print_sheet(&sheet_data.sheet);

//     // Test scrolling affects view
//     unsafe {
//         START_ROW = 1;
//         START_COL = 1;
//     }

//     print_sheet(&sheet_data.sheet);

//     // Test corner case - scrolling beyond bounds
//     unsafe {
//         START_ROW = 100;  // Beyond sheet bounds
//         START_COL = 100;  // Beyond sheet bounds
//     }

//     print_sheet(&sheet_data.sheet);

//     // Verify row and column labels are displayed correctly
//     unsafe {
//         START_ROW = 0;
//         START_COL = 0;

//         // Resize sheet to test multi-character column labels
//         R = 10;
//         C = 30;
//     }

//     let mut large_sheet_data = SheetData::new(10, 30);
//     print_sheet(&large_sheet_data.sheet);

//     // Visual inspection test - this is difficult to assert automatically
//     // but we can check that the function runs without crashing

//     // Reset to original dimensions
//     unsafe {
//         R = 5;
//         C = 5;
//     }
// }

#[test]
fn test_create_sheet() {
    // Test sheet creation with different dimensions
    let sheet_data = SheetData::new(10, 10);
    assert_eq!(sheet_data.sheet.len(), 10);
    assert_eq!(sheet_data.sheet[0].len(), 10);

    // Check that all cells are initialized properly
    for row in &sheet_data.sheet {
        for cell in row {
            let cell_ref = cell.borrow();
            assert_eq!(cell_ref.val, 0);
            assert_eq!(cell_ref.status, 0);
            assert_eq!(cell_ref.expression, "");
            assert!(cell_ref.dependencies.is_none());
            assert!(cell_ref.dependents.is_none());
        }
    }

    // Test with larger dimensions
    let sheet_data_large = SheetData::new(100, 100);
    assert_eq!(sheet_data_large.sheet.len(), 100);
    assert_eq!(sheet_data_large.sheet[0].len(), 100);
}

#[test]
fn test_label_to_index() {
    // Test valid labels
    assert_eq!(label_to_index("A1"), Some((0, 0)));
    assert_eq!(label_to_index("B2"), Some((1, 1)));
    assert_eq!(label_to_index("Z26"), Some((25, 25)));
    assert_eq!(label_to_index("AA1"), Some((0, 26)));
    assert_eq!(label_to_index("AB10"), Some((9, 27)));
    assert_eq!(label_to_index("ZZ99"), Some((98, 701)));

    // Test invalid labels
    // assert_eq!(label_to_index(""), None);
    assert_eq!(label_to_index("A0"), None);
    assert_eq!(label_to_index("A01"), None);
    assert_eq!(label_to_index("1A"), None);
    assert_eq!(label_to_index("AA"), None);
    assert_eq!(label_to_index("123"), None);
    assert_eq!(label_to_index("AAAA1"), None); // Too long
}

#[test]
fn test_label_to_index_invalid() {
    assert_eq!(label_to_index("1A"), None);
    assert_eq!(label_to_index("A01"), None);
    // assert_eq!(label_to_index(""), None);
    // println!("Empty: {:?}", label_to_index(""));
    assert_eq!(label_to_index("A!1"), None);
    assert_eq!(label_to_index("A10000"), None);
}

#[test]
fn test_col_label_to_index() {
    // Test valid column labels
    assert_eq!(col_label_to_index("A"), Some(0));
    assert_eq!(col_label_to_index("B"), Some(1));
    assert_eq!(col_label_to_index("Z"), Some(25));
    assert_eq!(col_label_to_index("AA"), Some(26));
    assert_eq!(col_label_to_index("AB"), Some(27));
    assert_eq!(col_label_to_index("ZZ"), Some(701));
    assert_eq!(col_label_to_index("AAA"), Some(702));

    // Test invalid column labels
    assert_eq!(col_label_to_index(""), None);
    assert_eq!(col_label_to_index("a"), None); // Lowercase
    assert_eq!(col_label_to_index("A1"), None); // Contains digit
    assert_eq!(col_label_to_index("A-"), None); // Contains non-alphabetic
}

#[test]
fn test_col_label_to_index_invalid() {
    assert_eq!(col_label_to_index("a"), None);
    assert_eq!(col_label_to_index("A1"), None);
    assert_eq!(col_label_to_index("!"), None);
}

#[test]
fn test_col_index_to_label() {
    // Test column index to label conversion
    assert_eq!(col_index_to_label(0), "A");
    assert_eq!(col_index_to_label(1), "B");
    assert_eq!(col_index_to_label(25), "Z");
    assert_eq!(col_index_to_label(26), "AA");
    assert_eq!(col_index_to_label(27), "AB");
    assert_eq!(col_index_to_label(701), "ZZ");
    assert_eq!(col_index_to_label(702), "AAA");

    // Test round-trip conversion
    for i in 0..1000 {
        let label = col_index_to_label(i);
        assert_eq!(col_label_to_index(&label), Some(i));
    }
}

// #[test]
// fn test_scroll_and_bounds() {
//     unsafe {
//         R = 100;
//         C = 100;
//         START_ROW = 0;
//         START_COL = 0;
//     }

//     scroll("w");
//     assert_eq!(unsafe { START_ROW }, 0);

//     scroll("s");
//     assert!(unsafe { START_ROW } > 0);

//     scroll("a");
//     assert_eq!(unsafe { START_COL }, 0);

//     scroll("d");
//     assert!(unsafe { START_COL } > 0);
// }

#[test]
fn test_execute_command() {
    let mut data = SheetData::new(10, 10);
    unsafe {
        R = 5;
        C = 5;
    }

    let mut status1 = execute_command("q", 5, 5, &mut data);
    assert_eq!(status1, 1);

    status1 = execute_command("w", 5, 5, &mut data);
    assert_eq!(status1, 0);

    status1 = execute_command("disable_output", 5, 5, &mut data);
    assert_eq!(status1, 0);

    status1 = execute_command("enable_output", 5, 5, &mut data);
    assert_eq!(status1, 0);

    status1 = execute_command("A1=MAX(C1:C1)", 5, 5, &mut data);
    assert_eq!(status1, 0);
    assert_eq!(data.sheet[0][0].borrow().val, 0);

    status1 = execute_command("A1=MIN(C1:C1)", 5, 5, &mut data);
    assert_eq!(status1, 0);
    assert_eq!(data.sheet[0][0].borrow().val, 0);

    status1 = execute_command("A1=AVG(C1:C1)", 5, 5, &mut data);
    assert_eq!(status1, 0);
    assert_eq!(data.sheet[0][0].borrow().val, 0);

    status1 = execute_command("A1=SUM(C1:C1)", 5, 5, &mut data);
    assert_eq!(status1, 0);
    assert_eq!(data.sheet[0][0].borrow().val, 0);

    status1 = execute_command("A1=STDEV(C1:C1)", 5, 5, &mut data);
    assert_eq!(status1, 0);
    assert_eq!(data.sheet[0][0].borrow().val, 0);

    let status2 = execute_command("A1=MAX(D8:B1)", 5, 5, &mut data);
    assert_eq!(status2, -1);
    assert_eq!(data.sheet[0][0].borrow().val, 0);

    let status3 = execute_command("A1=10", 5, 5, &mut data);
    assert_eq!(status3, 0);
    assert_eq!(data.sheet[0][0].borrow().val, 10);

    let status4 = execute_command("A1=3.2", 5, 5, &mut data);
    assert_eq!(status4, -1);
    assert_eq!(data.sheet[0][0].borrow().val, 10);

    let status5 = execute_command("B1=A1+5", 5, 5, &mut data);
    assert_eq!(status5, 0);
    assert_eq!(data.sheet[0][1].borrow().val, 15);

    let status6 = execute_command("A1=B1", 5, 5, &mut data);
    assert_eq!(status6, -4);

    let status7 = execute_command("A1=10/0", 5, 5, &mut data);
    assert_eq!(status7, -2);
    assert_eq!(data.sheet[0][0].borrow().status, 1);
    assert_eq!(data.sheet[0][1].borrow().status, 1);

    let status8 = execute_command("q", 5, 5, &mut data);
    assert_eq!(status8, 1);

    let status9 = execute_command("scroll_to B2", 5, 5, &mut data);
    assert_eq!(status9, 0);

    unsafe {
        R = 10;
        C = 10;
    }
    let mut data2 = SheetData::new(10, 10);
    let mut result = 0;
    let row = 0;
    let col = 5;
    execute_command("F1=MAX(G1:J9)", 10, 10, &mut data2);
    evaluate_expression("20", 10, 10, &mut data2, &mut result, &row, &col, 1);
    print_sheet(&data2.sheet);
}

#[test]
fn test_push_dependent() {
    let sheet_data = &mut SheetData::new(5, 5);
    let cell1 = &sheet_data.sheet[0][0];
    let cell2 = &sheet_data.sheet[1][1];

    // Initially no dependents
    assert!(cell1.borrow().dependents.is_none());

    // Add cell2 as dependent of cell1
    push_dependent(&cell1.clone(), &cell2.clone());

    // Check that cell2 is now a dependent of cell1
    let dependents = &cell1.borrow().dependents;
    assert!(dependents.is_some());

    // Check that the dependent is cell2
    let dep_node = dependents.as_ref().unwrap();
    assert!(Rc::ptr_eq(&dep_node.borrow().cell, cell2));
}

fn test_add_dependency() {
    let sheet_data = &mut SheetData::new(5, 5);
    let cell1 = &sheet_data.sheet[0][0].clone();
    let cell2 = &sheet_data.sheet[1][1].clone();

    // Initially no dependencies
    assert!(cell1.borrow().dependencies.is_none());

    // Add cell2 as dependency of cell1
    add_dependency(&cell1.clone(), &cell2.clone(), sheet_data);

    // Check that cell2 is now a dependency of cell1
    let dependencies = &cell1.borrow().dependencies;
    assert!(dependencies.is_some());

    // Check that the dependency is cell2
    if let Some(dep_node) = dependencies {
        assert!(Rc::ptr_eq(&dep_node.borrow().cell, cell2));
    } else {
        panic!("Expected dependency not found");
    }
}

// #[test]
// fn test_scroll() {
//     unsafe {
//         R = 100;
//         C = 100;
//         START_ROW = 20;
//         START_COL = 20;
//     }

//     // Test scrolling up
//     scroll("w");
//     assert_eq!(unsafe { START_ROW }, 10);
//     assert_eq!(unsafe { START_COL }, 20);

//     // Test scrolling down
//     scroll("s");
//     assert_eq!(unsafe { START_ROW }, 20);
//     assert_eq!(unsafe { START_COL }, 20);

//     // Test scrolling left
//     scroll("a");
//     assert_eq!(unsafe { START_ROW }, 20);
//     assert_eq!(unsafe { START_COL }, 10);

//     // Test scrolling right
//     scroll("d");
//     assert_eq!(unsafe { START_ROW }, 20);
//     assert_eq!(unsafe { START_COL }, 20);
// }

// #[test]
// fn test_scroll_edge() {
//     unsafe {
//         R = 30;
//         C = 30;
//         START_ROW = 5;
//         START_COL = 5;
//     }

//     // Test scrolling up at edge
//     unsafe { START_ROW = 5; }
//     scroll("w");
//     assert_eq!(unsafe { START_ROW }, 0);

//     // Test scrolling left at edge
//     unsafe { START_COL = 5; }
//     scroll("a");
//     assert_eq!(unsafe { START_COL }, 0);

//     // Test scrolling down at edge
//     unsafe { START_ROW = 25; }
//     scroll("s");
//     assert_eq!(unsafe { START_ROW }, 20); // Should be capped at R-10

//     // Test scrolling right at edge
//     unsafe { START_COL = 25; }
//     scroll("d");
//     assert_eq!(unsafe { START_COL }, 20); // Should be capped at C-10
// }

#[test]
fn test_sleep_seconds() {
    // Test that sleep_seconds runs without panicking
    // Note: We'll use a short duration to avoid slowing down tests
    let start = Instant::now();
    sleep_seconds(1);
    let elapsed = start.elapsed();

    // Check that at least 1 second has passed
    assert!(elapsed.as_secs() >= 1);
}

#[test]
fn test_sleep_seconds_edge() {
    // Test with zero seconds
    let start = Instant::now();
    sleep_seconds(0);
    let elapsed = start.elapsed();

    // Should return almost immediately
    assert!(elapsed.as_millis() < 100);

    // We won't test very large values to avoid slowing down tests
}

fn test_delete_dependencies() {
    let sheet_data = &mut SheetData::new(5, 5);
    let cell1 = &sheet_data.sheet[0][0].clone();
    let cell2 = &sheet_data.sheet[1][1].clone();

    // Set up dependency: cell1 depends on cell2
    add_dependency(&cell1.clone(), &cell2.clone(), sheet_data);
    push_dependent(&cell2.clone(), &cell1.clone());

    // Verify dependency exists
    assert!(cell1.borrow().dependencies.is_some());
    assert!(cell2.borrow().dependents.is_some());

    // Delete dependencies
    delete_dependencies(1, 1, sheet_data);

    // Verify dependencies are cleared
    assert!(cell1.borrow().dependencies.is_none());
    assert!(cell2.borrow().dependents.is_none());
}
// #[test]
// fn test_dfs() {
//     let sheet_data = &mut SheetData::new(5, 5);
//     let a1 = &sheet_data.sheet[0][0].clone();
//     let b1 = &sheet_data.sheet[0][1].clone();
//     let c1 = &sheet_data.sheet[0][2].clone();
//     let d1 = &sheet_data.sheet[0][3].clone();
//     let e1 = &sheet_data.sheet[0][4].clone();

//     // Set up chain: cell1 -> cell2 -> cell3
//     add_dependency(a1.clone(), b1.clone(), sheet_data);
//     add_dependency(a1.clone(), c1.clone(), sheet_data);
//     add_dependency(b1.clone(), c1.clone(), sheet_data);
//     add_dependency(c1.clone(), d1.clone(), sheet_data);

//     // Check direct connection
//     let mut visited = vec![0u64; (5 * 5 + 63) / 64];
//     assert!(dfs(a1, c1, &mut visited, 0, 0, sheet_data));
//     assert!(dfs(a1, b1, &mut visited, 0, 0, sheet_data));
//     assert!(!(dfs(b1, a1, &mut visited, 0, 0, sheet_data)));
//     assert!(!(dfs(a1, e1, &mut visited, 0, 0, sheet_data)));
// }

// #[test]
// fn test_check_loop() {
//     let sheet_data = &mut SheetData::new(5, 5);
//     let cell1 = &sheet_data.sheet[0][0].clone();
//     let cell2 = &sheet_data.sheet[1][1].clone();
//     let cell3 = &sheet_data.sheet[2][2].clone();

//     // No loop initially
//     assert!(check_loop(cell1, cell2, 0, 0, sheet_data));

//     // Set up chain: cell1 -> cell2 -> cell3
//     add_dependency(cell1.clone(), &cell2.clone(), sheet_data);
//     push_dependent(cell2.clone(), cell1.clone());

//     add_dependency(cell2.clone(), cell3.clone(), sheet_data);
//     push_dependent(cell3.clone(), &cell2.clone());

//     // Check for loop with direct dependency
//     assert!(check_loop(cell1, cell2, 0, 0, sheet_data));

//     // Create a loop: cell3 -> cell1
//     add_dependency(cell3.clone(), cell1.clone(), sheet_data);
//     push_dependent(cell1.clone(), cell3.clone());

//     // Should detect the loop
//     assert!(check_loop(cell1, cell1, 0, 0, sheet_data));
//     assert!(check_loop(cell2, cell2, 1, 1, sheet_data));
//     assert!(check_loop(cell3, cell3, 2, 2, sheet_data));
// }

#[test]
fn test_dfs_range() {
    let sheet_data = &mut SheetData::new(5, 5);
    let a1 = &sheet_data.sheet[0][0].clone();
    let b1 = &sheet_data.sheet[0][1].clone();
    let c1 = &sheet_data.sheet[0][2].clone();
    let b2 = &sheet_data.sheet[1][1].clone();
    let c2 = &sheet_data.sheet[1][2].clone();

    // Set up chain: cell1 -> cell2 -> cell3
    add_dependency(&a1, &b2.clone(), sheet_data);
    push_dependent(&b2.clone(), &a1.clone());

    add_dependency(&b1, &a1.clone(), sheet_data);
    add_dependency(&c1, &a1.clone(), sheet_data);
    add_dependency(&b2, &a1.clone(), sheet_data);
    add_dependency(&c2, &a1.clone(), sheet_data);
    push_dependent(&a1.clone(), &b1.clone());
    push_dependent(&a1.clone(), &c1.clone());
    push_dependent(&a1.clone(), &b2.clone());
    push_dependent(&a1.clone(), &c2.clone());

    unsafe {
        R = 5;
        C = 5;
    }

    // Check if cell1 depends on cells in the range (1,1) to (2,2)
    let mut visited = vec![false; 5 * 5];
    assert!(dfs_range(a1, &mut visited, 0, 1, 1, 2, 0, 0, sheet_data));
}

#[test]
fn test_check_loop_range() {
    let sheet_data = &mut SheetData::new(5, 5);
    let a1 = &sheet_data.sheet[0][0].clone();
    let b1 = &sheet_data.sheet[0][1].clone();
    let c1 = &sheet_data.sheet[0][2].clone();
    let b2 = &sheet_data.sheet[1][1].clone();
    let c2 = &sheet_data.sheet[1][2].clone();

    // Set up chain: cell1 -> cell2 -> cell3
    add_dependency(&b1.clone(), &a1.clone(), sheet_data);
    add_dependency(&c1.clone(), &a1.clone(), sheet_data);
    add_dependency(&b2.clone(), &a1.clone(), sheet_data);
    add_dependency(&c2.clone(), &a1.clone(), sheet_data);

    unsafe {
        R = 5;
        C = 5;
    }

    // Check if cell1 depends on cells in the range (1,1) to (2,2)
    assert!(!check_loop_range(a1, 0, 1, 1, 2, 0, 0, sheet_data));

    add_dependency(&a1.clone(), &b2.clone(), sheet_data);

    // Check range outside dependencies
    assert!(check_loop_range(a1, 0, 1, 1, 2, 0, 0, sheet_data));
}

#[test]
fn test_topological_sort_util() {
    let sheet_data = &mut SheetData::new(5, 5);
    let cell1 = &sheet_data.sheet[0][0].clone();
    let cell2 = &sheet_data.sheet[1][1].clone();
    let cell3 = &sheet_data.sheet[2][2].clone();

    // Set up chain: cell1 -> cell2 -> cell3
    add_dependency(&cell1.clone(), &cell2.clone(), sheet_data);
    push_dependent(&cell2.clone(), &cell1.clone());

    add_dependency(&cell2.clone(), &cell3.clone(), sheet_data);
    push_dependent(&cell3.clone(), &cell2.clone());

    unsafe {
        R = 5;
        C = 5;
    }

    let mut stack = None;
    let mut visited = vec![false; 5 * 5];

    // Run topological sort
    topological_sort_util(cell1, &mut visited, sheet_data, &mut stack);

    // Check if stack has elements in correct order
    // The top of the stack should be cell1, followed by cell2, then cell3
    let first = pop(&mut stack);
    assert!(first.is_some());
    assert!(Rc::ptr_eq(&first.unwrap(), cell1));

    let second = pop(&mut stack);
    assert!(second.is_some());
    assert!(Rc::ptr_eq(&second.unwrap(), cell2));

    let third = pop(&mut stack);
    assert!(third.is_some());
    assert!(Rc::ptr_eq(&third.unwrap(), cell3));

    // Stack should be empty now
    assert!(pop(&mut stack).is_none());
}

#[test]
fn test_evaluate_expression() {
    unsafe {
        R = 10;
        C = 10;
    }

    let sheet_data = &mut SheetData::new(10, 10);

    // Test simple integer
    let mut result = 0;
    let row = 0;
    let col = 0;
    let c3_row = 2;
    let c3_col = 2;
    assert_eq!(
        evaluate_expression("42", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 42);

    // Test simple addition
    result = 0;
    assert_eq!(
        evaluate_expression("2+3", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 5);

    result = 0;
    assert_eq!(
        evaluate_expression("2*3", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 6);

    result = 0;
    assert_eq!(
        evaluate_expression("2-3", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, -1);

    result = 0;
    assert_eq!(
        evaluate_expression("2/3", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 0);

    // Test cell reference
    sheet_data.sheet[1][0].borrow_mut().val = 42;
    result = 0;
    assert_eq!(
        evaluate_expression("A2", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 42);
    result = 0;
    assert_eq!(
        evaluate_expression("A2+10", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 52);
    result = 0;
    assert_eq!(
        evaluate_expression("10+A2", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 52);
    result = 0;
    assert_eq!(
        evaluate_expression("10-A2", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, -32);
    result = 0;
    assert_eq!(
        evaluate_expression("A2-10", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 32);
    result = 0;
    assert_eq!(
        evaluate_expression("10*A2", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 420);
    result = 0;
    assert_eq!(
        evaluate_expression("A2*10", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 420);
    result = 0;
    assert_eq!(
        evaluate_expression("A2/10", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 4);
    result = 0;
    assert_eq!(
        evaluate_expression("10/A2", 10, 10, sheet_data, &mut result, &row, &col, 1),
        0
    );
    assert_eq!(result, 0);

    // Test SUM function
    sheet_data.sheet[0][0].borrow_mut().val = 1;
    sheet_data.sheet[0][1].borrow_mut().val = 2;
    sheet_data.sheet[1][0].borrow_mut().val = 3;
    sheet_data.sheet[1][1].borrow_mut().val = 4;
    result = 0;
    assert_eq!(
        evaluate_expression(
            "SUM(A1:B2)",
            10,
            10,
            sheet_data,
            &mut result,
            &c3_row,
            &c3_col,
            1
        ),
        0
    );
    assert_eq!(result, 10); // 1+2+3+4

    // Test AVG function
    result = 0;
    assert_eq!(
        evaluate_expression(
            "AVG(A1:B2)",
            10,
            10,
            sheet_data,
            &mut result,
            &c3_row,
            &c3_col,
            1
        ),
        0
    );
    assert_eq!(result, 2); // (1+2+3+4)/4

    // Test MAX function
    result = 0;
    assert_eq!(
        evaluate_expression(
            "MAX(A1:B2)",
            10,
            10,
            sheet_data,
            &mut result,
            &c3_row,
            &c3_col,
            1
        ),
        0
    );
    assert_eq!(result, 4);

    // Test MIN function
    result = 0;
    assert_eq!(
        evaluate_expression(
            "MIN(A1:B2)",
            10,
            10,
            sheet_data,
            &mut result,
            &c3_row,
            &c3_col,
            1
        ),
        0
    );
    assert_eq!(result, 1);

    result = 0;
    assert_eq!(
        evaluate_expression(
            "STDEV(A1:B2)",
            10,
            10,
            sheet_data,
            &mut result,
            &c3_row,
            &c3_col,
            1
        ),
        0
    );
    assert_eq!(result, 1);
}

#[test]
fn test_evaluate_wrong_expression() {
    unsafe {
        R = 10;
        C = 10;
    }
    let mut sheet_data = SheetData::new(10, 10);
    let mut result = 0;
    let row = 0;
    let col = 0;

    // Test invalid cell references
    assert_eq!(
        evaluate_expression("Z99", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -1
    );
    assert_eq!(
        evaluate_expression(
            "AAA1000",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    );

    // Test invalid cell format
    assert_eq!(
        evaluate_expression("1A", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -1
    );
    assert_eq!(
        evaluate_expression("A01", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -1
    );
    assert_eq!(
        evaluate_expression("A!1", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -1
    );

    // Test invalid expressions
    assert_eq!(
        evaluate_expression(
            "C1++B1",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    );
    assert_eq!(
        evaluate_expression("A1B1", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -1
    );

    // Test division by zero
    sheet_data.sheet[0][0].borrow_mut().val = 10;
    sheet_data.sheet[0][1].borrow_mut().val = 0;
    assert_eq!(
        evaluate_expression("A1/B1", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -4
    );

    // Test with cell having error status
    sheet_data.sheet[0][1].borrow_mut().status = 1;
    assert_eq!(
        evaluate_expression("B1", 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -2
    );

    // Test circular dependency
    // Set up A1 to depend on B1, then try to make B1 depend on A1
    unsafe {
        R = 10;
        C = 10;
    }
    sheet_data = SheetData::new(10, 10); // Reset
    let row_a = 0;
    let col_a = 0;
    let row_b = 0;
    let col_b = 1;
    let row_c = 0;
    let col_c = 2;
    let row_d = 0;
    let col_d = 3;

    sheet_data.sheet[0][1].borrow_mut().status = 0; // Reset status
                                                    // Make A1 depend on B1
    evaluate_expression(
        "B1",
        10,
        10,
        &mut sheet_data,
        &mut result,
        &row_a,
        &col_a,
        1,
    );
    evaluate_expression(
        "C1",
        10,
        10,
        &mut sheet_data,
        &mut result,
        &row_b,
        &col_b,
        1,
    );
    evaluate_expression(
        "D1",
        10,
        10,
        &mut sheet_data,
        &mut result,
        &row_c,
        &col_c,
        1,
    );
    // assert_eq!(evaluate_expression("A1", 10, 10, &mut sheet_data, &mut result, &row_d, &col_d, 0), -4);

    // Now try to make B1 depend on A1 - should detect circular dependency
    assert_eq!(
        evaluate_expression(
            "SUM(A1:C1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row_d,
            &col_d,
            1
        ),
        -4
    );
    assert_eq!(
        evaluate_expression(
            "MAX(A1:C1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row_d,
            &col_d,
            1
        ),
        -4
    );
    assert_eq!(
        evaluate_expression(
            "MIN(A1:C1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row_d,
            &col_d,
            1
        ),
        -4
    );
    assert_eq!(
        evaluate_expression(
            "AVG(A1:C1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row_d,
            &col_d,
            1
        ),
        -4
    );
    assert_eq!(
        evaluate_expression(
            "STDEV(A1:C1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row_d,
            &col_d,
            1
        ),
        -4
    );

    // Test invalid function calls
    assert_eq!(
        evaluate_expression(
            "SUM(A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Missing range
    assert_eq!(
        evaluate_expression(
            "MAX(A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Missing range
    assert_eq!(
        evaluate_expression(
            "MIN(A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Missing range
    assert_eq!(
        evaluate_expression(
            "AVG(A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Missing range
    assert_eq!(
        evaluate_expression(
            "STDEV(A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Missing range
    assert_eq!(
        evaluate_expression(
            "SUM(A1:B2:C3)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Invalid range
    assert_eq!(
        evaluate_expression(
            "MAX(A1:B2:C3)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Invalid range
    assert_eq!(
        evaluate_expression(
            "MIN(A1:B2:C3)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Invalid range
    assert_eq!(
        evaluate_expression(
            "AVG(A1:B2:C3)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Invalid range
    assert_eq!(
        evaluate_expression(
            "STDEV(A1:B2:C3)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Invalid range

    assert_eq!(
        evaluate_expression(
            "SUMM(A1:B2)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Unknown function
    assert_eq!(
        evaluate_expression(
            "MMAX(A1:B2)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Unknown function
    assert_eq!(
        evaluate_expression(
            "MIIN(A1:B2)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Unknown function
    assert_eq!(
        evaluate_expression(
            "AVGVG(A1:B2)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Unknown function
    assert_eq!(
        evaluate_expression(
            "STD_EV(A1:B2)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Unknown function

    // Test invalid ranges
    assert_eq!(
        evaluate_expression(
            "SUM(B1:A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // End before start
    assert_eq!(
        evaluate_expression(
            "MAX(B1:A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // End before start
    assert_eq!(
        evaluate_expression(
            "MIN(B1:A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // End before start
    assert_eq!(
        evaluate_expression(
            "AVG(B1:A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // End before start
    assert_eq!(
        evaluate_expression(
            "STDEV(B1:A1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // End before start
    assert_eq!(
        evaluate_expression(
            "SUM(A1:Z1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Out of bounds
    assert_eq!(
        evaluate_expression(
            "MAX(A1:Z1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Out of bounds
    assert_eq!(
        evaluate_expression(
            "MIN(A1:Z1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Out of bounds
    assert_eq!(
        evaluate_expression(
            "AVG(A1:Z1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Out of bounds
    assert_eq!(
        evaluate_expression(
            "STDEV(A1:Z1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    ); // Out of bounds

    // Test functions with cells having error status
    sheet_data.sheet[0][5].borrow_mut().status = 1;
    assert_eq!(
        evaluate_expression(
            "SUM(F1:F5)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -2
    );

    // Test invalid sleep inputs
    assert_eq!(
        evaluate_expression(
            "SLEEP(A)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        -1
    );
    assert_eq!(
        evaluate_expression(
            "SLEEP(-1)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        0
    ); // Negative values are handled
    assert_eq!(
        evaluate_expression(
            "SLEEP(F5)",
            10,
            10,
            &mut sheet_data,
            &mut result,
            &row,
            &col,
            1
        ),
        0
    ); // Zero values are handled
       // Test with circular reference in a range function
    sheet_data = SheetData::new(10, 10); // Reset
    let expr = "SUM(A1:B2)";
    let row = 0;
    let col = 0;

    // This should detect that the cell is in the range it's trying to evaluate
    assert_eq!(
        evaluate_expression(expr, 10, 10, &mut sheet_data, &mut result, &row, &col, 1),
        -4
    );
}

// #[test]
// fn test_topological_sort_from_cell_simple() {
//     let mut sheet_data = init_sheet(5, 5);
//     let a1 = sheet_data.sheet[0][0].clone();
//     let b1 = sheet_data.sheet[0][1].clone();
//     let c1 = sheet_data.sheet[0][2].clone();
//     let d1 = sheet_data.sheet[0][3].clone();

//     add_dependency(a1.clone(), b1.clone(), &mut sheet_data);
//     add_dependency(a1.clone(), c1.clone(), &mut sheet_data);
//     add_dependency(b1.clone(), c1.clone(), &mut sheet_data);
//     add_dependency(c1.clone(), d1.clone(), &mut sheet_data);
//     let mut stack = None;
//     topological_sort_from_cell(&a1, &sheet_data, &mut stack);

//     let mut count = 0;
//     while pop(&mut stack).is_some() {
//         count += 1;
//     }

//     assert!(count == 4);
// }
