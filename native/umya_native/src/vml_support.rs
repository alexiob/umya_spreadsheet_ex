use crate::atoms;
use crate::UmyaSpreadsheet; // Correct path to UmyaSpreadsheet
use rustler::types::atom::Atom;
use rustler::{NifResult, ResourceArc}; // Removed unused Env
use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use umya_spreadsheet::structs::vml::Shape;
use umya_spreadsheet::structs::{EmbeddedObjectProperties, OleObject};

// Helper function to convert string shape ID to numeric ID
fn string_to_numeric_shape_id(shape_id: &str) -> u32 {
    // First try to parse as a number directly
    if let Ok(num) = shape_id.parse::<u32>() {
        return num;
    }

    // If not a number, create a hash-based numeric ID
    let mut hasher = DefaultHasher::new();
    shape_id.hash(&mut hasher);
    let hash = hasher.finish();
    // Use the lower 31 bits to ensure we get a positive u32
    (hash & 0x7FFFFFFF) as u32
}

// Helper function to find an OLE object by shape ID
fn find_ole_object_by_shape_id<'a>(
    worksheet: &'a mut umya_spreadsheet::Worksheet,
    shape_id: &str,
) -> Option<&'a mut OleObject> {
    let shape_id_num = string_to_numeric_shape_id(shape_id);

    worksheet
        .get_ole_objects_mut()
        .get_ole_object_mut()
        .iter_mut()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        })
}

// VML Shape NIF functions

#[rustler::nif]
fn create_vml_shape(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Convert string shape_id to numeric ID for embedded object properties
    let shape_id_num = string_to_numeric_shape_id(&shape_id);

    // Create a new VML shape
    let mut vml_shape = Shape::default();
    vml_shape.set_type("rect"); // Default to rectangle
    vml_shape.set_filled(true);
    vml_shape.set_fill_color("#FFFFFF"); // Default white fill
    vml_shape.set_stroked(true);
    vml_shape.set_stroke_color("#000000"); // Default black stroke
    vml_shape.set_stroke_weight("1pt"); // Default stroke weight

    // Create a new OLE object to hold the VML shape
    let mut ole_object = OleObject::default();
    ole_object.set_shape(vml_shape);

    // Create embedded object properties with the shape ID
    let mut properties = EmbeddedObjectProperties::default();
    properties.set_shape_id(shape_id_num);
    ole_object.set_embedded_object_properties(properties);

    // Add the OLE object to the worksheet's OLE objects collection
    worksheet.get_ole_objects_mut().set_ole_object(ole_object);

    let result = Ok(atoms::ok());

    // Explicitly drop the guard to ensure mutex is released before returning
    drop(guard);

    result
}

#[rustler::nif]
fn set_vml_shape_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    style: String,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate style parameter
    if style.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Style cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape style
    ole_object.get_shape_mut().set_style(&style);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

#[rustler::nif]
fn set_vml_shape_type(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    shape_type: String,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Verify the shape type is valid
    let valid_types = ["rect", "oval", "line", "polyline", "roundrect", "arc"];
    if !valid_types.contains(&shape_type.as_str()) {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            format!(
                "Invalid shape type '{}'. Supported types: {:?}",
                shape_type, valid_types
            ),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape type
    ole_object.get_shape_mut().set_type(&shape_type);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

#[rustler::nif]
fn set_vml_shape_filled(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    filled: bool,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape filled property
    ole_object.get_shape_mut().set_filled(filled);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

#[rustler::nif]
fn set_vml_shape_fill_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    fill_color: String,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Validate fill color (check for valid color format)
    if !fill_color.starts_with('#') && !fill_color.starts_with("rgb(") {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Invalid fill color format. Expected '#RRGGBB' or 'rgb(r,g,b)'".to_string(),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape fill color
    ole_object.get_shape_mut().set_fill_color(&fill_color);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

#[rustler::nif]
fn set_vml_shape_stroked(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    stroked: bool,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape stroked property
    ole_object.get_shape_mut().set_stroked(stroked);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

#[rustler::nif]
fn set_vml_shape_stroke_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    stroke_color: String,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Validate stroke color (check for valid color format)
    if !stroke_color.starts_with('#') && !stroke_color.starts_with("rgb(") {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Invalid stroke color format. Expected '#RRGGBB' or 'rgb(r,g,b)'".to_string(),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape stroke color
    ole_object.get_shape_mut().set_stroke_color(&stroke_color);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

#[rustler::nif]
fn set_vml_shape_stroke_weight(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
    stroke_weight: String,
) -> NifResult<Atom> {
    let mut guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Validate stroke weight (should be a number with optional unit like "pt")
    if stroke_weight.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Stroke weight cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name and get mutable reference
    let worksheet = match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let ole_object = match find_ole_object_by_shape_id(worksheet, &shape_id) {
        Some(obj) => obj,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))));
        }
    };

    // Update the VML shape stroke weight
    ole_object.get_shape_mut().set_stroke_weight(&stroke_weight);

    let result = Ok(atoms::ok());
    drop(guard);
    result
}

// VML Shape getter functions
#[rustler::nif]
fn get_vml_shape_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, String)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let style = obj.get_shape().get_style().to_string();
            drop(guard);
            Ok((atoms::ok(), style))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}

#[rustler::nif]
fn get_vml_shape_type(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, String)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let shape_type = obj.get_shape().get_type().to_string();
            drop(guard);
            Ok((atoms::ok(), shape_type))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}

#[rustler::nif]
fn get_vml_shape_filled(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, bool)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let filled = *obj.get_shape().get_filled();
            drop(guard);
            Ok((atoms::ok(), filled))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}

#[rustler::nif]
fn get_vml_shape_fill_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, String)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let fill_color = obj.get_shape().get_fill_color().to_string();
            drop(guard);
            Ok((atoms::ok(), fill_color))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}

#[rustler::nif]
fn get_vml_shape_stroked(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, bool)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let stroked = *obj.get_shape().get_stroked();
            drop(guard);
            Ok((atoms::ok(), stroked))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}

#[rustler::nif]
fn get_vml_shape_stroke_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, String)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let stroke_color = obj.get_shape().get_stroke_color().to_string();
            drop(guard);
            Ok((atoms::ok(), stroke_color))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}

#[rustler::nif]
fn get_vml_shape_stroke_weight(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    shape_id: String,
) -> NifResult<(Atom, String)> {
    let guard = match resource.spreadsheet.lock() {
        Ok(guard) => guard,
        Err(_) => {
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                "Failed to lock spreadsheet resource".to_string(),
            ))))
        }
    };

    // Validate shape ID
    if shape_id.is_empty() {
        drop(guard);
        return Err(rustler::Error::Term(Box::new((
            atoms::error(),
            "Shape ID cannot be empty".to_string(),
        ))));
    }

    // Find the worksheet by name
    let worksheet = match guard.get_sheet_by_name(&sheet_name) {
        Some(ws) => ws,
        None => {
            drop(guard);
            return Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Worksheet '{}' not found", sheet_name),
            ))));
        }
    };

    // Find the OLE object with the specified shape ID
    let shape_id_num = string_to_numeric_shape_id(&shape_id);
    let ole_object = worksheet
        .get_ole_objects()
        .get_ole_object()
        .iter()
        .find(|ole_object| {
            ole_object.get_embedded_object_properties().get_shape_id() == &shape_id_num
        });

    match ole_object {
        Some(obj) => {
            let stroke_weight = obj.get_shape().get_stroke_weight().to_string();
            drop(guard);
            Ok((atoms::ok(), stroke_weight))
        }
        None => {
            drop(guard);
            Err(rustler::Error::Term(Box::new((
                atoms::error(),
                format!("Shape with ID '{}' not found", shape_id),
            ))))
        }
    }
}
