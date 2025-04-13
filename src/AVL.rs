use std::cell::RefCell;
use std::cmp::max;
use std::rc::Rc;

pub type Link = Option<Rc<RefCell<AvlNode>>>;

#[derive(Clone)]
pub struct Cell {
    pub value: i32,
}

pub struct AvlNode {
    pub cell: Rc<RefCell<Cell>>,
    pub left: Link,
    pub right: Link,
    pub height: i32,
}

impl AvlNode {
    pub fn new(cell: Rc<RefCell<Cell>>) -> Rc<RefCell<Self>> {
        Rc::new(RefCell::new(Self {
            cell,
            left: None,
            right: None,
            height: 1,
        }))
    }
}

fn height(node: &Link) -> i32 {
    node.as_ref().map_or(0, |n| n.borrow().height)
}

fn get_balance(node: &Rc<RefCell<AvlNode>>) -> i32 {
    height(&node.borrow().left) - height(&node.borrow().right)
}

fn rotate_right(y: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let x = y.borrow_mut().left.take().unwrap();
    let t2 = x.borrow_mut().right.take();

    y.borrow_mut().left = t2;
    x.borrow_mut().right = Some(y.clone());

    y.borrow_mut().height = max(height(&y.borrow().left), height(&y.borrow().right)) + 1;
    x.borrow_mut().height = max(height(&x.borrow().left), height(&x.borrow().right)) + 1;
    x
}

fn rotate_left(x: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let y = x.borrow_mut().right.take().unwrap();
    let t2 = y.borrow_mut().left.take();

    x.borrow_mut().right = t2;
    y.borrow_mut().left = Some(x.clone());

    x.borrow_mut().height = max(height(&x.borrow().left), height(&x.borrow().right)) + 1;
    y.borrow_mut().height = max(height(&y.borrow().left), height(&y.borrow().right)) + 1;
    y
}

pub fn insert(node: Link, cell: Rc<RefCell<Cell>>) -> Link {
    if let Some(n) = node {
        let mut n_borrow = n.borrow_mut();

        if cell.borrow().value < n_borrow.cell.borrow().value {
            n_borrow.left = insert(n_borrow.left.clone(), cell.clone());
        } else if cell.borrow().value > n_borrow.cell.borrow().value {
            n_borrow.right = insert(n_borrow.right.clone(), cell.clone());
        } else {
            return Some(n.clone()); // Duplicate values not allowed
        }

        n_borrow.height = 1 + max(height(&n_borrow.left), height(&n_borrow.right));
        drop(n_borrow);

        let balance = get_balance(&n);
        let node = n.clone();

        if balance > 1 && cell.borrow().value < node.borrow().left.as_ref().unwrap().borrow().cell.borrow().value {
            return Some(rotate_right(node));
        }
        if balance < -1 && cell.borrow().value > node.borrow().right.as_ref().unwrap().borrow().cell.borrow().value {
            return Some(rotate_left(node));
        }
        if balance > 1 && cell.borrow().value > node.borrow().left.as_ref().unwrap().borrow().cell.borrow().value {
            let left_rotated = rotate_left(node.borrow().left.clone().unwrap());
            node.borrow_mut().left = Some(left_rotated);
            return Some(rotate_right(node));
        }
        if balance < -1 && cell.borrow().value < node.borrow().right.as_ref().unwrap().borrow().cell.borrow().value {
            let right_rotated = rotate_right(node.borrow().right.clone().unwrap());
            node.borrow_mut().right = Some(right_rotated);
            return Some(rotate_left(node));
        }

        Some(node)
    } else {
        Some(AvlNode::new(cell))
    }
}

pub fn find(node: &Link, value: i32) -> Link {
    match node {
        Some(n) => {
            let n_borrow = n.borrow();
            if value == n_borrow.cell.borrow().value {
                Some(n.clone())
            } else if value < n_borrow.cell.borrow().value {
                find(&n_borrow.left, value)
            } else {
                find(&n_borrow.right, value)
            }
        }
        None => None,
    }
}

fn min_value_node(node: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let mut current = node;
    loop {
        let left = {
            let curr_borrow = current.borrow();
            curr_borrow.left.clone()
        };

        if let Some(left_node) = left {
            current = left_node;
        } else {
            break;
        }
    }
    current
}

pub fn delete_node(root: Link, value: i32) -> Link {
    if let Some(node) = root {
        let mut node_borrow = node.borrow_mut();

        if value < node_borrow.cell.borrow().value {
            node_borrow.left = delete_node(node_borrow.left.clone(), value);
        } else if value > node_borrow.cell.borrow().value {
            node_borrow.right = delete_node(node_borrow.right.clone(), value);
        } else {
            // Node to be deleted found
            if node_borrow.left.is_none() || node_borrow.right.is_none() {
                return node_borrow.left.clone().or(node_borrow.right.clone());
            } else {
                let temp = min_value_node(node_borrow.right.clone().unwrap());
                let replacement_value = temp.borrow().cell.borrow().value;
                node_borrow.cell = Rc::new(RefCell::new(Cell { value: replacement_value }));
                node_borrow.right = delete_node(node_borrow.right.clone(), replacement_value);
            }
        }

        node_borrow.height = 1 + max(height(&node_borrow.left), height(&node_borrow.right));
        drop(node_borrow);

        let balance = get_balance(&node);

        if balance > 1 && get_balance(&node.borrow().left.as_ref().unwrap()) >= 0 {
            return Some(rotate_right(node));
        }

        if balance > 1 && get_balance(&node.borrow().left.as_ref().unwrap()) < 0 {
            let left_rotated = rotate_left(node.borrow().left.clone().unwrap());
            node.borrow_mut().left = Some(left_rotated);
            return Some(rotate_right(node));
        }

        if balance < -1 && get_balance(&node.borrow().right.as_ref().unwrap()) <= 0 {
            return Some(rotate_left(node));
        }

        if balance < -1 && get_balance(&node.borrow().right.as_ref().unwrap()) > 0 {
            let right_rotated = rotate_right(node.borrow().right.clone().unwrap());
            node.borrow_mut().right = Some(right_rotated);
            return Some(rotate_left(node));
        }

        Some(node)
    } else {
        None
    }
}

pub struct AvlTree {
    root: Link,
}

impl AvlTree {
    pub fn new() -> Self {
        AvlTree { root: None }
    }

    pub fn insert(&mut self, value: i32) {
        let cell = Rc::new(RefCell::new(Cell { value }));
        self.root = insert(self.root.take(), cell);
    }

    pub fn delete(&mut self, value: i32) {
        self.root = delete_node(self.root.take(), value);
    }

    pub fn find(&self, value: i32) -> bool {
        find(&self.root, value).is_some()
    }

    pub fn inorder(&self) {
        fn traverse(node: &Link) {
            if let Some(n) = node {
                traverse(&n.borrow().left);
                print!("{} ", n.borrow().cell.borrow().value);
                traverse(&n.borrow().right);
            }
        }
        traverse(&self.root);
        println!();
    }
}