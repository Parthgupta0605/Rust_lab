mod AVL;
use AVL::AvlTree;

fn main() {
    let mut tree = AvlTree::new();

    // Insert values
    for &value in &[30, 20, 40, 10, 25, 35, 50] {
        println!("Inserting {}", value);
        tree.insert(value);
    }

    println!("In-order traversal after insertions:");
    tree.inorder(); // should print sorted values

    // Search
    println!("Searching for 25: {}", tree.find(25));
    println!("Searching for 99: {}", tree.find(99));

    // Deletion
    println!("Deleting 20");
    tree.delete(20);
    println!("In-order traversal after deleting 20:");
    tree.inorder();

    println!("Deleting 30");
    tree.delete(30);
    println!("In-order traversal after deleting 30:");
    tree.inorder();
}
