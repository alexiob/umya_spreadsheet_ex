use rustler::{NifResult, ResourceArc};
use std::sync::Mutex;
use umya_spreadsheet::{EmbeddedObjectProperties, OleObject, OleObjects};

// Import the main spreadsheet struct from the parent module
use crate::UmyaSpreadsheet;

/// Resource for OleObjects
pub struct OleObjectsResource {
    pub ole_objects: Mutex<OleObjects>,
}

/// Resource for OleObject
pub struct OleObjectResource {
    pub ole_object: Mutex<OleObject>,
}

/// Resource for EmbeddedObjectProperties
pub struct EmbeddedObjectPropertiesResource {
    pub properties: Mutex<EmbeddedObjectProperties>,
}

/// Create a new OleObjects collection
#[rustler::nif]
pub fn create_ole_objects() -> NifResult<ResourceArc<OleObjectsResource>> {
    let ole_objects = OleObjects::default();
    let resource = ResourceArc::new(OleObjectsResource {
        ole_objects: Mutex::new(ole_objects),
    });
    Ok(resource)
}

/// Create a new OleObject
#[rustler::nif]
pub fn create_ole_object() -> NifResult<ResourceArc<OleObjectResource>> {
    let ole_object = OleObject::default();
    let resource = ResourceArc::new(OleObjectResource {
        ole_object: Mutex::new(ole_object),
    });
    Ok(resource)
}

/// Create a new EmbeddedObjectProperties
#[rustler::nif]
pub fn create_embedded_object_properties(
) -> NifResult<ResourceArc<EmbeddedObjectPropertiesResource>> {
    let properties = EmbeddedObjectProperties::default();
    let resource = ResourceArc::new(EmbeddedObjectPropertiesResource {
        properties: Mutex::new(properties),
    });
    Ok(resource)
}

/// Get OLE objects from a worksheet
#[rustler::nif]
pub fn get_ole_objects(
    spreadsheet_res: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<ResourceArc<OleObjectsResource>> {
    let spreadsheet = spreadsheet_res.spreadsheet.lock().unwrap();
    let worksheet = spreadsheet
        .get_sheet_by_name(&sheet_name)
        .ok_or_else(|| rustler::Error::Term(Box::new("Sheet not found")))?;

    let ole_objects = worksheet.get_ole_objects().clone();
    let resource = ResourceArc::new(OleObjectsResource {
        ole_objects: Mutex::new(ole_objects),
    });
    Ok(resource)
}

/// Set OLE objects to a worksheet
#[rustler::nif]
pub fn set_ole_objects(
    spreadsheet_res: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    ole_objects_res: ResourceArc<OleObjectsResource>,
) -> NifResult<()> {
    let mut spreadsheet = spreadsheet_res.spreadsheet.lock().unwrap();
    let ole_objects = ole_objects_res.ole_objects.lock().unwrap();

    let worksheet = spreadsheet
        .get_sheet_by_name_mut(&sheet_name)
        .ok_or_else(|| rustler::Error::Term(Box::new("Sheet not found")))?;

    worksheet.set_ole_objects(ole_objects.clone());
    Ok(())
}

/// Add an OLE object to a collection
#[rustler::nif]
pub fn add_ole_object(
    ole_objects_res: ResourceArc<OleObjectsResource>,
    ole_object_res: ResourceArc<OleObjectResource>,
) -> NifResult<()> {
    let mut ole_objects = ole_objects_res.ole_objects.lock().unwrap();
    let ole_object = ole_object_res.ole_object.lock().unwrap();

    ole_objects.set_ole_object(ole_object.clone());
    Ok(())
}

/// Get all OLE objects from a collection
#[rustler::nif]
pub fn get_ole_object_list(
    ole_objects_res: ResourceArc<OleObjectsResource>,
) -> NifResult<Vec<ResourceArc<OleObjectResource>>> {
    let ole_objects = ole_objects_res.ole_objects.lock().unwrap();
    let mut results = Vec::new();

    for ole_object in ole_objects.get_ole_object() {
        let resource = ResourceArc::new(OleObjectResource {
            ole_object: Mutex::new(ole_object.clone()),
        });
        results.push(resource);
    }

    Ok(results)
}

/// Get count of OLE objects in a collection
#[rustler::nif]
pub fn count_ole_objects(ole_objects_res: ResourceArc<OleObjectsResource>) -> NifResult<usize> {
    let ole_objects = ole_objects_res.ole_objects.lock().unwrap();
    Ok(ole_objects.get_ole_object().len())
}

/// Check if collection has any OLE objects
#[rustler::nif]
pub fn has_ole_objects(ole_objects_res: ResourceArc<OleObjectsResource>) -> NifResult<bool> {
    let ole_objects = ole_objects_res.ole_objects.lock().unwrap();
    Ok(!ole_objects.get_ole_object().is_empty())
}

/// Get requires property from OLE object
#[rustler::nif]
pub fn get_ole_object_requires(
    ole_object_res: ResourceArc<OleObjectResource>,
) -> NifResult<String> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    Ok(ole_object.get_requires().to_string())
}

/// Set requires property for OLE object
#[rustler::nif]
pub fn set_ole_object_requires(
    ole_object_res: ResourceArc<OleObjectResource>,
    requires: String,
) -> NifResult<()> {
    let mut ole_object = ole_object_res.ole_object.lock().unwrap();
    ole_object.set_requires(requires);
    Ok(())
}

/// Get program ID from OLE object
#[rustler::nif]
pub fn get_ole_object_prog_id(ole_object_res: ResourceArc<OleObjectResource>) -> NifResult<String> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    Ok(ole_object.get_prog_id().to_string())
}

/// Set program ID for OLE object
#[rustler::nif]
pub fn set_ole_object_prog_id(
    ole_object_res: ResourceArc<OleObjectResource>,
    prog_id: String,
) -> NifResult<()> {
    let mut ole_object = ole_object_res.ole_object.lock().unwrap();
    ole_object.set_prog_id(prog_id);
    Ok(())
}

/// Get object extension from OLE object
#[rustler::nif]
pub fn get_ole_object_extension(
    ole_object_res: ResourceArc<OleObjectResource>,
) -> NifResult<String> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    Ok(ole_object.get_object_extension().to_string())
}

/// Set object extension for OLE object
#[rustler::nif]
pub fn set_ole_object_extension(
    ole_object_res: ResourceArc<OleObjectResource>,
    extension: String,
) -> NifResult<()> {
    let mut ole_object = ole_object_res.ole_object.lock().unwrap();
    ole_object.set_object_extension(extension);
    Ok(())
}

/// Get object data from OLE object
#[rustler::nif]
pub fn get_ole_object_data(
    ole_object_res: ResourceArc<OleObjectResource>,
) -> NifResult<Option<Vec<u8>>> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    Ok(ole_object.get_object_data().map(|data| data.to_vec()))
}

/// Set object data for OLE object
#[rustler::nif]
pub fn set_ole_object_data(
    ole_object_res: ResourceArc<OleObjectResource>,
    data: Vec<u8>,
) -> NifResult<()> {
    let mut ole_object = ole_object_res.ole_object.lock().unwrap();
    ole_object.set_object_data(data);
    Ok(())
}

/// Get embedded object properties from OLE object
#[rustler::nif]
pub fn get_ole_object_properties(
    ole_object_res: ResourceArc<OleObjectResource>,
) -> NifResult<ResourceArc<EmbeddedObjectPropertiesResource>> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    let properties = ole_object.get_embedded_object_properties().clone();
    let resource = ResourceArc::new(EmbeddedObjectPropertiesResource {
        properties: Mutex::new(properties),
    });
    Ok(resource)
}

/// Set embedded object properties for OLE object
#[rustler::nif]
pub fn set_ole_object_properties(
    ole_object_res: ResourceArc<OleObjectResource>,
    properties_res: ResourceArc<EmbeddedObjectPropertiesResource>,
) -> NifResult<()> {
    let mut ole_object = ole_object_res.ole_object.lock().unwrap();
    let properties = properties_res.properties.lock().unwrap();
    ole_object.set_embedded_object_properties(properties.clone());
    Ok(())
}

/// Get program ID from embedded object properties
#[rustler::nif]
pub fn get_embedded_object_prog_id(
    properties_res: ResourceArc<EmbeddedObjectPropertiesResource>,
) -> NifResult<String> {
    let properties = properties_res.properties.lock().unwrap();
    Ok(properties.get_prog_id().to_string())
}

/// Set program ID for embedded object properties
#[rustler::nif]
pub fn set_embedded_object_prog_id(
    properties_res: ResourceArc<EmbeddedObjectPropertiesResource>,
    prog_id: String,
) -> NifResult<()> {
    let mut properties = properties_res.properties.lock().unwrap();
    properties.set_prog_id(prog_id);
    Ok(())
}

/// Get shape ID from embedded object properties
#[rustler::nif]
pub fn get_embedded_object_shape_id(
    properties_res: ResourceArc<EmbeddedObjectPropertiesResource>,
) -> NifResult<u32> {
    let properties = properties_res.properties.lock().unwrap();
    Ok(*properties.get_shape_id())
}

/// Set shape ID for embedded object properties
#[rustler::nif]
pub fn set_embedded_object_shape_id(
    properties_res: ResourceArc<EmbeddedObjectPropertiesResource>,
    shape_id: u32,
) -> NifResult<()> {
    let mut properties = properties_res.properties.lock().unwrap();
    properties.set_shape_id(shape_id);
    Ok(())
}

/// Load OLE object from file
#[rustler::nif]
pub fn load_ole_object_from_file(
    file_path: String,
    prog_id: String,
) -> NifResult<ResourceArc<OleObjectResource>> {
    use std::fs;
    use std::path::Path;

    let path = Path::new(&file_path);
    if !path.exists() {
        return Err(rustler::Error::Term(Box::new("File not found")));
    }

    let data = fs::read(path).map_err(|_| rustler::Error::Term(Box::new("Failed to read file")))?;

    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_string();

    let mut ole_object = OleObject::default();
    ole_object.set_prog_id(prog_id);
    ole_object.set_object_extension(extension);
    ole_object.set_object_data(data);

    let resource = ResourceArc::new(OleObjectResource {
        ole_object: Mutex::new(ole_object),
    });
    Ok(resource)
}

/// Save OLE object data to file
#[rustler::nif]
pub fn save_ole_object_to_file(
    ole_object_res: ResourceArc<OleObjectResource>,
    file_path: String,
) -> NifResult<()> {
    use std::fs;

    let ole_object = ole_object_res.ole_object.lock().unwrap();
    if let Some(data) = ole_object.get_object_data() {
        fs::write(file_path, data)
            .map_err(|_| rustler::Error::Term(Box::new("Failed to write file")))?;
        Ok(())
    } else {
        Err(rustler::Error::Term(Box::new("No object data to save")))
    }
}

/// Create OLE object with file data
#[rustler::nif]
pub fn create_ole_object_with_data(
    prog_id: String,
    extension: String,
    data: Vec<u8>,
) -> NifResult<ResourceArc<OleObjectResource>> {
    let mut ole_object = OleObject::default();
    ole_object.set_prog_id(prog_id);
    ole_object.set_object_extension(extension);
    ole_object.set_object_data(data);

    let resource = ResourceArc::new(OleObjectResource {
        ole_object: Mutex::new(ole_object),
    });
    Ok(resource)
}

/// Check if OLE object is binary format
#[rustler::nif]
pub fn is_ole_object_binary(ole_object_res: ResourceArc<OleObjectResource>) -> NifResult<bool> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    Ok(ole_object.get_object_extension() == "bin")
}

/// Check if OLE object is Excel format
#[rustler::nif]
pub fn is_ole_object_excel(ole_object_res: ResourceArc<OleObjectResource>) -> NifResult<bool> {
    let ole_object = ole_object_res.ole_object.lock().unwrap();
    Ok(ole_object.get_object_extension() == "xlsx")
}
