use std::cell::RefCell;
use std::cmp::max;
use std::rc::Rc;

type Link = Option<Rc<RefCell<AvlNode>>>;

#[derive(Clone)]
pub struct Cell {
    // Placeholder field; update based on your actual `cell.h`
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
    {
        let mut y_mut = y.borrow_mut();
        y_mut.left = t2;
    }
    x.borrow_mut().right = Some(y.clone());

    y.borrow_mut().height = max(height(&y.borrow().left), height(&y.borrow().right)) + 1;
    x.borrow_mut().height = max(height(&x.borrow().left), height(&x.borrow().right)) + 1;
    x
}

fn rotate_left(x: Rc<RefCell<AvlNode>>) -> Rc<RefCell<AvlNode>> {
    let y = x.borrow_mut().right.take().unwrap();
    let t2 = y.borrow_mut().left.take();
    {
        let mut x_mut = x.borrow_mut();
        x_mut.right = t2;
    }
    y.borrow_mut().left = Some(x.clone());

    x.borrow_mut().height = max(height(&x.borrow().left), height(&x.borrow().right)) + 1;
    y.borrow_mut().height = max(height(&y.borrow().left), height(&y.borrow().right)) + 1;
    y
}

fn insert(node: Link, cell: Rc<RefCell<Cell>>) -> Link {
    if let Some(n) = node {
        let mut n_borrow = n.borrow_mut();

        if cell.borrow().value < n_borrow.cell.borrow().value {
            n_borrow.left = insert(n_borrow.left.clone(), cell);
        } else if cell.borrow().value > n_borrow.cell.borrow().value {
            n_borrow.right = insert(n_borrow.right.clone(), cell);
        } else {
            return Some(n.clone());
        }

        n_borrow.height = 1 + max(height(&n_borrow.left), height(&n_borrow.right));
        drop(n_borrow); // Drop borrow before rotation

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

fn find(node: &Link, value: i32) -> Link {
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
    while let Some(left) = current.borrow().left.clone() {
        current = left;
    }
    current
}

fn delete_node(root: Link, value: i32) -> Link {
    if let Some(node) = root {
        let mut node_borrow = node.borrow_mut();

        if value < node_borrow.cell.borrow().value {
            node_borrow.left = delete_node(node_borrow.left.clone(), value);
        } else if value > node_borrow.cell.borrow().value {
            node_borrow.right = delete_node(node_borrow.right.clone(), value);
        } else {
            if node_borrow.left.is_none() || node_borrow.right.is_none() {
                return node_borrow.left.clone().or(node_borrow.right.clone());
            } else {
                let temp = min_value_node(node_borrow.right.clone().unwrap());
                node_borrow.cell = temp.borrow().cell.clone();
                node_borrow.right = delete_node(node_borrow.right.clone(), temp.borrow().cell.borrow().value);
            }
        }

        node_borrow.height = 1 + max(height(&node_borrow.left), height(&node_borrow.right));
        drop(node_borrow); // Drop borrow before rotation

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
            let right_rotated
        }
::contentReference[oaicite:3]{index=3}
    }
}
 
