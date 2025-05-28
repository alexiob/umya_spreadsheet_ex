// Document Properties functions for UmyaSpreadsheet
use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use umya_spreadsheet;

use crate::atoms;
use crate::UmyaSpreadsheet;

/// Get a custom document property value by name
#[rustler::nif]
pub fn get_custom_property(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
) -> NifResult<(Atom, String)> {
    let guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties().get_custom_properties();

    for property in custom_properties.get_custom_document_property_list() {
        if property.get_name() == name {
            let value = property.get_value();
            return Ok((atoms::ok(), value.to_string()));
        }
    }

    Err(NifError::Term(Box::new((
        atoms::not_found(),
        "Property not found".to_string(),
    ))))
}

/// Set a custom document property with string value
#[rustler::nif]
pub fn set_custom_property_string(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
    value: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties_mut().get_custom_properties_mut();

    // Check if property already exists and update it
    for property in custom_properties.get_custom_document_property_list_mut() {
        if property.get_name() == name {
            property.set_value_string(value);
            return Ok(atoms::ok());
        }
    }

    // Create new property if it doesn't exist
    let mut new_property =
        umya_spreadsheet::structs::custom_properties::CustomDocumentProperty::default();
    new_property.set_name(name);
    new_property.set_value_string(value);
    custom_properties.add_custom_document_property_list(new_property);

    Ok(atoms::ok())
}

/// Set a custom document property with numeric value
#[rustler::nif]
pub fn set_custom_property_number(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
    value: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties_mut().get_custom_properties_mut();

    // Check if property already exists and update it
    for property in custom_properties.get_custom_document_property_list_mut() {
        if property.get_name() == name {
            property.set_value_number(value);
            return Ok(atoms::ok());
        }
    }

    // Create new property if it doesn't exist
    let mut new_property =
        umya_spreadsheet::structs::custom_properties::CustomDocumentProperty::default();
    new_property.set_name(name);
    new_property.set_value_number(value);
    custom_properties.add_custom_document_property_list(new_property);

    Ok(atoms::ok())
}

/// Set a custom document property with boolean value
#[rustler::nif]
pub fn set_custom_property_bool(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
    value: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties_mut().get_custom_properties_mut();

    // Check if property already exists and update it
    for property in custom_properties.get_custom_document_property_list_mut() {
        if property.get_name() == name {
            property.set_value_bool(value);
            return Ok(atoms::ok());
        }
    }

    // Create new property if it doesn't exist
    let mut new_property =
        umya_spreadsheet::structs::custom_properties::CustomDocumentProperty::default();
    new_property.set_name(name);
    new_property.set_value_bool(value);
    custom_properties.add_custom_document_property_list(new_property);

    Ok(atoms::ok())
}

/// Set a custom document property with date value
#[rustler::nif]
pub fn set_custom_property_date(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
    year: i32,
    month: i32,
    day: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties_mut().get_custom_properties_mut();

    // Check if property already exists and update it
    for property in custom_properties.get_custom_document_property_list_mut() {
        if property.get_name() == name {
            property.set_value_date(year, month, day);
            return Ok(atoms::ok());
        }
    }

    // Create new property if it doesn't exist
    let mut new_property =
        umya_spreadsheet::structs::custom_properties::CustomDocumentProperty::default();
    new_property.set_name(name);
    new_property.set_value_date(year, month, day);
    custom_properties.add_custom_document_property_list(new_property);

    Ok(atoms::ok())
}

/// Remove a custom document property by name
#[rustler::nif]
pub fn remove_custom_property(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties_mut().get_custom_properties_mut();
    let property_list = custom_properties.get_custom_document_property_list_mut();

    // Find and remove the property
    let mut index_to_remove = None;
    for (index, property) in property_list.iter().enumerate() {
        if property.get_name() == name {
            index_to_remove = Some(index);
            break;
        }
    }

    if let Some(index) = index_to_remove {
        property_list.remove(index);
        Ok(atoms::ok())
    } else {
        Err(NifError::Term(Box::new((
            atoms::not_found(),
            "Property not found".to_string(),
        ))))
    }
}

/// Get all custom document property names
#[rustler::nif]
pub fn get_custom_property_names(
    resource: ResourceArc<UmyaSpreadsheet>,
) -> NifResult<(Atom, Vec<String>)> {
    let guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties().get_custom_properties();
    let mut names = Vec::new();

    for property in custom_properties.get_custom_document_property_list() {
        names.push(property.get_name().to_string());
    }

    Ok((atoms::ok(), names))
}

/// Check if a custom document property exists
#[rustler::nif]
pub fn has_custom_property(
    resource: ResourceArc<UmyaSpreadsheet>,
    name: String,
) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties().get_custom_properties();

    for property in custom_properties.get_custom_document_property_list() {
        if property.get_name() == name {
            return Ok(true);
        }
    }

    Ok(false)
}

/// Get the number of custom document properties
#[rustler::nif]
pub fn get_custom_properties_count(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<usize> {
    let guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties().get_custom_properties();
    let count = custom_properties.get_custom_document_property_list().len();

    Ok(count)
}

/// Clear all custom document properties
#[rustler::nif]
pub fn clear_custom_properties(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    let custom_properties = guard.get_properties_mut().get_custom_properties_mut();
    custom_properties
        .get_custom_document_property_list_mut()
        .clear();

    Ok(atoms::ok())
}

// Core document properties functions

/// Get the document title
#[rustler::nif]
pub fn get_title(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_title().to_string())
}

/// Set the document title
#[rustler::nif]
pub fn set_title(resource: ResourceArc<UmyaSpreadsheet>, title: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_title(title);
    Ok(atoms::ok())
}

/// Get the document description
#[rustler::nif]
pub fn get_description(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_description().to_string())
}

/// Set the document description
#[rustler::nif]
pub fn set_description(
    resource: ResourceArc<UmyaSpreadsheet>,
    description: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_description(description);
    Ok(atoms::ok())
}

/// Get the document subject
#[rustler::nif]
pub fn get_subject(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_subject().to_string())
}

/// Set the document subject
#[rustler::nif]
pub fn set_subject(resource: ResourceArc<UmyaSpreadsheet>, subject: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_subject(subject);
    Ok(atoms::ok())
}

/// Get the document keywords
#[rustler::nif]
pub fn get_keywords(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_keywords().to_string())
}

/// Set the document keywords
#[rustler::nif]
pub fn set_keywords(resource: ResourceArc<UmyaSpreadsheet>, keywords: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_keywords(keywords);
    Ok(atoms::ok())
}

/// Get the document creator
#[rustler::nif]
pub fn get_creator(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_creator().to_string())
}

/// Set the document creator
#[rustler::nif]
pub fn set_creator(resource: ResourceArc<UmyaSpreadsheet>, creator: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_creator(creator);
    Ok(atoms::ok())
}

/// Get the document last modified by
#[rustler::nif]
pub fn get_last_modified_by(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_last_modified_by().to_string())
}

/// Set the document last modified by
#[rustler::nif]
pub fn set_last_modified_by(
    resource: ResourceArc<UmyaSpreadsheet>,
    last_modified_by: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard
        .get_properties_mut()
        .set_last_modified_by(last_modified_by);
    Ok(atoms::ok())
}

/// Get the document category
#[rustler::nif]
pub fn get_category(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_category().to_string())
}

/// Set the document category
#[rustler::nif]
pub fn set_category(resource: ResourceArc<UmyaSpreadsheet>, category: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_category(category);
    Ok(atoms::ok())
}

/// Get the document company
#[rustler::nif]
pub fn get_company(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_company().to_string())
}

/// Set the document company
#[rustler::nif]
pub fn set_company(resource: ResourceArc<UmyaSpreadsheet>, company: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_company(company);
    Ok(atoms::ok())
}

/// Get the document manager
#[rustler::nif]
pub fn get_manager(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_manager().to_string())
}

/// Set the document manager
#[rustler::nif]
pub fn set_manager(resource: ResourceArc<UmyaSpreadsheet>, manager: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_manager(manager);
    Ok(atoms::ok())
}

/// Get the document created date
#[rustler::nif]
pub fn get_created(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_created().to_string())
}

/// Set the document created date
#[rustler::nif]
pub fn set_created(resource: ResourceArc<UmyaSpreadsheet>, created: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_created(created);
    Ok(atoms::ok())
}

/// Get the document modified date
#[rustler::nif]
pub fn get_modified(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_properties().get_modified().to_string())
}

/// Set the document modified date
#[rustler::nif]
pub fn set_modified(resource: ResourceArc<UmyaSpreadsheet>, modified: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();
    guard.get_properties_mut().set_modified(modified);
    Ok(atoms::ok())
}
