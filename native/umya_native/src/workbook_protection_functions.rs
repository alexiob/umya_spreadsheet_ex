use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Encoder, Env, Term};
use std::collections::HashMap;
use std::panic::{self, AssertUnwindSafe};

#[rustler::nif]
pub fn is_workbook_protected(
    env: Env,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Check if workbook protection exists
        let is_protected = spreadsheet.get_workbook_protection().is_some();

        (atoms::ok(), is_protected).encode(env)
    }));

    match result {
        Ok(term) => term,
        Err(_) => (atoms::error(), "Error occurred in is_workbook_protected").encode(env),
    }
}

#[rustler::nif]
pub fn get_workbook_protection_details(
    env: Env,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Check if workbook protection exists
        match spreadsheet.get_workbook_protection() {
            Some(protection) => {
                let mut details = HashMap::new();

                // Add lock status properties
                details.insert(
                    "lock_structure".to_string(),
                    protection.get_lock_structure().to_string(),
                );
                details.insert(
                    "lock_windows".to_string(),
                    protection.get_lock_windows().to_string(),
                );
                details.insert(
                    "lock_revision".to_string(),
                    protection.get_lock_revision().to_string(),
                );

                // We don't expose password hashes for security reasons

                (atoms::ok(), details).encode(env)
            }
            None => (atoms::error(), "Workbook is not protected").encode(env),
        }
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_workbook_protection_details",
        )
            .encode(env),
    }
}
