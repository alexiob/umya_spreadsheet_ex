// Error helper functions for umya_native
// This module previously contained error handling utilities that have been replaced
// with direct NifError::Term usage throughout the codebase for better error handling.

#[allow(dead_code)]
mod legacy {
    use rustler::{Atom, Encoder, Error as NifError, NifResult, Term};

    /// Legacy error handler - replaced with direct NifError::Term usage
    pub fn handle_error<T: AsRef<str>>(_message: T) -> NifResult<Atom> {
        Err(NifError::Term(Box::new(crate::atoms::error())))
    }

    /// Create an error tuple with a reason string for Elixir {:error, reason} pattern
    pub fn create_error_tuple<'a>(env: rustler::Env<'a>, reason: &str) -> Term<'a> {
        let error_atom = crate::atoms::error().encode(env);
        let reason_term = reason.encode(env);
        (error_atom, reason_term).encode(env)
    }

    /// Create an ok tuple with a value for Elixir {:ok, value} pattern
    pub fn create_ok_tuple<'a, T: rustler::Encoder>(env: rustler::Env<'a>, value: T) -> Term<'a> {
        let ok_atom = crate::atoms::ok().encode(env);
        let value_term = value.encode(env);
        (ok_atom, value_term).encode(env)
    }

    /// Return an error tuple term directly (for functions that return NifResult<Term>)
    pub fn error_tuple<'a>(env: rustler::Env<'a>, reason: &str) -> NifResult<Term<'a>> {
        Ok(create_error_tuple(env, reason))
    }

    /// Return an ok tuple term directly (for functions that return NifResult<Term>)
    pub fn ok_tuple<'a, T: rustler::Encoder>(
        env: rustler::Env<'a>,
        value: T,
    ) -> NifResult<Term<'a>> {
        Ok(create_ok_tuple(env, value))
    }
}
