use rustler::{Atom, Error as NifError, NifResult};

pub fn handle_error<T: AsRef<str>>(message: T) -> NifResult<Atom> {
    let error_str = message.as_ref();
    println!("Error in Rust: {}", error_str);
    Err(NifError::Term(Box::new(crate::atoms::error())))
}
