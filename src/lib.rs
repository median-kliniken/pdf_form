mod utils;

use derive_error::Error;
use lopdf::content::{Content, Operation};
use lopdf::dictionary;
use lopdf::{Dictionary, Document, Object, ObjectId, StringFormat};
use std::collections::VecDeque;
use std::io;
use std::io::Write;
use std::path::Path;
use std::str;
use std::str::from_utf8;

use crate::utils::*;

/// A PDF Form that contains fillable fields
///
/// Use this struct to load an existing PDF with a fillable form using the `load` method.  It will
/// analyze the PDF and identify the fields. Then you can get and set the content of the fields by
/// index.
pub struct Form {
    doc: Document,
    form_ids: Vec<ObjectId>,
}

/// The possible types of fillable form fields in a PDF
#[derive(Debug, Eq, PartialEq)]
pub enum FieldType {
    Button,
    Radio,
    CheckBox,
    ListBox,
    ComboBox,
    Text,
    Unknown,
}

#[derive(Debug, Error)]
/// Errors that may occur while loading a PDF
pub enum LoadError {
    /// An Lopdf Error
    LopdfError(lopdf::Error),
    /// The reference `ObjectId` did not point to any values
    #[error(non_std, no_from)]
    NoSuchReference(ObjectId),
    /// An element that was expected to be a reference was not a reference
    NotAReference,
}

/// Errors That may occur while setting values in a form
#[derive(Debug, Error)]
pub enum ValueError {
    /// The method used to set the state is incompatible with the type of the field
    TypeMismatch,
    /// One or more selected values are not valid choices
    InvalidSelection,
    /// Multiple values were selected when only one was allowed
    TooManySelected,
    /// Readonly field cannot be edited
    Readonly,
    /// A provided combobox value was not a valid choice
    UnknownOption,
}

#[derive(Debug)]
pub struct SingleCheckboxState {
    pub widget_id: ObjectId,
    pub on_value: String,
    pub is_checked: bool,
    pub readonly: bool,
    pub required: bool,
}

/// The current state of a form field
#[derive(Debug)]
pub enum FieldState {
    /// Push buttons have no state
    Button,
    /// `selected` is the singular option from `options` that is selected
    Radio {
        selected: String,
        options: Vec<String>,
        readonly: bool,
        required: bool,
    },
    /// A group of check boxes
    CheckBoxGroup { states: Vec<SingleCheckboxState> },
    /// The toggle state of the checkbox
    CheckBox {
        is_checked: bool,
        readonly: bool,
        required: bool,
    },
    /// `selected` is the list of selected options from `options`
    ListBox {
        selected: Vec<String>,
        options: Vec<String>,
        multiselect: bool,
        readonly: bool,
        required: bool,
    },
    /// `selected` is the list of selected options from `options`
    ComboBox {
        selected: Vec<String>,
        options: Vec<String>,
        editable: bool,
        readonly: bool,
        required: bool,
    },
    /// User Text Input
    Text {
        text: String,
        readonly: bool,
        required: bool,
    },
    /// Unknown fields have no state
    Unknown,
}

trait PdfObjectDeref {
    fn deref<'a>(&self, doc: &'a Document) -> Result<&'a Object, LoadError>;
}

impl PdfObjectDeref for Object {
    fn deref<'a>(&self, doc: &'a Document) -> Result<&'a Object, LoadError> {
        match *self {
            Object::Reference(oid) => doc.objects.get(&oid).ok_or(LoadError::NoSuchReference(oid)),
            _ => Err(LoadError::NotAReference),
        }
    }
}

impl Form {
    /// Takes a reader containing a PDF with a fillable form, analyzes the content, and attempts to
    /// identify all of the fields the form has.
    pub fn load_from<R: io::Read>(reader: R) -> Result<Self, LoadError> {
        let doc = Document::load_from(reader)?;
        Self::load_doc(doc)
    }

    /// Takes a path to a PDF with a fillable form, analyzes the file, and attempts to identify all
    /// of the fields the form has.
    pub fn load<P: AsRef<Path>>(path: P) -> Result<Self, LoadError> {
        let doc = Document::load(path)?;
        Self::load_doc(doc)
    }

    fn load_doc(mut doc: Document) -> Result<Self, LoadError> {
        let mut form_ids = Vec::new();
        let mut queue = VecDeque::new();
        // Block so borrow of doc ends before doc is moved into the result
        {
            doc.decompress();

            let acroform = doc
                .objects
                .get_mut(
                    &doc.trailer
                        .get(b"Root")?
                        .deref(&doc)?
                        .as_dict()?
                        .get(b"AcroForm")?
                        .as_reference()?,
                )
                .ok_or(LoadError::NotAReference)?
                .as_dict_mut()?;

            acroform.set("NeedAppearances", Object::Boolean(true));

            let fields_list = acroform.get(b"Fields")?.as_array()?;
            queue.append(&mut VecDeque::from(fields_list.clone()));

            // Iterate over the fields
            while let Some(objref) = queue.pop_front() {
                let obj = objref.deref(&doc)?;
                if let Object::Dictionary(ref dict) = *obj {
                    // If the field has FT, it actually takes input.  Save this
                    if dict.get(b"FT").is_ok() {
                        form_ids.push(objref.as_reference().unwrap());
                    }

                    // If this field has kids, they might have FT, so add them to the queue
                    if let Ok(&Object::Array(ref kids)) = dict.get(b"Kids") {
                        queue.append(&mut VecDeque::from(kids.clone()));
                    }
                }
            }
        }
        Ok(Form { doc, form_ids })
    }

    /// Returns the number of fields the form has
    pub fn len(&self) -> usize {
        self.form_ids.len()
    }

    /// Returns true if empty
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// Gets the type of field of the given index
    ///
    /// # Panics
    /// This function will panic if the index is greater than the number of fields
    pub fn get_type(&self, n: usize) -> FieldType {
        // unwraps should be fine because load should have verified everything exists
        let field = self
            .doc
            .objects
            .get(&self.form_ids[n])
            .unwrap()
            .as_dict()
            .unwrap();

        let type_str = field.get(b"FT").unwrap().as_name().unwrap();
        if type_str == b"Btn" {
            let flags = ButtonFlags::from_bits_truncate(get_field_flags(field));
            if flags.intersects(ButtonFlags::RADIO | ButtonFlags::NO_TOGGLE_TO_OFF) {
                FieldType::Radio
            } else if flags.intersects(ButtonFlags::PUSHBUTTON) {
                FieldType::Button
            } else {
                FieldType::CheckBox
            }
        } else if type_str == b"Ch" {
            let flags = ChoiceFlags::from_bits_truncate(get_field_flags(field));
            if flags.intersects(ChoiceFlags::COBMO) {
                FieldType::ComboBox
            } else {
                FieldType::ListBox
            }
        } else if type_str == b"Tx" {
            FieldType::Text
        } else {
            FieldType::Unknown
        }
    }

    /// Gets the name of field of the given index
    ///
    /// # Panics
    /// This function will panic if the index is greater than the number of fields
    pub fn get_name(&self, n: usize) -> Option<String> {
        // unwraps should be fine because load should have verified everything exists
        let field = self
            .doc
            .objects
            .get(&self.form_ids[n])
            .unwrap()
            .as_dict()
            .unwrap();

        // The "T" key refers to the name of the field
        match field.get(b"T") {
            Ok(Object::String(data, _)) => String::from_utf8(data.clone()).ok(),
            _ => None,
        }
    }

    /// Gets the types of all of the fields in the form
    pub fn get_all_types(&self) -> Vec<FieldType> {
        let mut res = Vec::with_capacity(self.len());
        for i in 0..self.len() {
            res.push(self.get_type(i))
        }
        res
    }

    /// Gets the names of all of the fields in the form
    pub fn get_all_names(&self) -> Vec<Option<String>> {
        let mut res = Vec::with_capacity(self.len());
        for i in 0..self.len() {
            res.push(self.get_name(i))
        }
        res
    }

    /// Gets the state of field of the given index
    ///
    /// # Panics
    /// This function will panic if the index is greater than the number of fields
    pub fn get_state(&self, n: usize) -> FieldState {
        let object = self.doc.objects.get(&self.form_ids[n]).unwrap();

        let field = object.as_dict().unwrap();
        match self.get_type(n) {
            FieldType::Button => FieldState::Button,
            FieldType::Radio => FieldState::Radio {
                selected: match field.get(b"V") {
                    Ok(name) => decode_pdf_string(name).unwrap_or_else(String::new),
                    _ => match field.get(b"AS") {
                        Ok(name) => decode_pdf_string(name).unwrap_or_else(String::new),
                        _ => "".to_owned(),
                    },
                },
                options: self.get_possibilities(self.form_ids[n]),
                readonly: is_read_only(field),
                required: is_required(field),
            },
            FieldType::CheckBox => {
                if let Ok(kids) = field.get(b"Kids").and_then(|o| o.as_array()) {
                    // Grouped checkboxes
                    let mut checkboxes = Vec::new();

                    for kid_ref in kids {
                        if let Ok(kid_id) = kid_ref.as_reference() {
                            if let Ok(widget_obj) = self.doc.get_object(kid_id) {
                                if let Ok(widget) = widget_obj.as_dict() {
                                    if let Ok(ap) = widget.get(b"AP").and_then(|o| o.as_dict()) {
                                        if let Ok(n_dict) = ap.get(b"N").and_then(|o| o.as_dict()) {
                                            for (name, _) in n_dict.iter() {
                                                if name != b"Off" {
                                                    let is_checked = widget
                                                        .get(b"AS")
                                                        .ok()
                                                        .and_then(|v| v.as_name().ok())
                                                        .map(|n| n == name)
                                                        .unwrap_or(false);
                                                    checkboxes.push(SingleCheckboxState {
                                                        widget_id: kid_id,
                                                        on_value: String::from_utf8_lossy(name)
                                                            .into(),
                                                        is_checked,
                                                        readonly: is_read_only(widget),
                                                        required: is_required(widget),
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    FieldState::CheckBoxGroup { states: checkboxes }
                } else {
                    // Simple checkbox
                    let is_checked = match field.get(b"V") {
                        Ok(name) => matches!(name.as_name().unwrap_or(b""), b"Yes" | b"On"),
                        _ => match field.get(b"AS") {
                            Ok(name) => matches!(name.as_name().unwrap_or(b""), b"Yes" | b"On"),
                            _ => false,
                        },
                    };

                    FieldState::CheckBox {
                        is_checked,
                        readonly: is_read_only(field),
                        required: is_required(field),
                    }
                }
            }
            FieldType::ListBox => {
                FieldState::ListBox {
                    // V field in a list box can be either text for one option, an array for many
                    // options, or null
                    selected: match field.get(b"V") {
                        Ok(selection) => match *selection {
                            Object::String(ref s, StringFormat::Literal) => {
                                vec![str::from_utf8(&s).unwrap().to_owned()]
                            }
                            Object::Array(ref chosen) => {
                                let mut res = Vec::new();
                                for obj in chosen {
                                    if let Some(string) = decode_pdf_string(obj) {
                                        res.push(string);
                                    }
                                }
                                res
                            }
                            _ => Vec::new(),
                        },
                        _ => Vec::new(),
                    },
                    // The options is an array of either text elements or arrays where the second
                    // element is what we want
                    options: match self.find_opt_in_kids_or_parents(&field) {
                        Some(&Object::Array(ref options)) => options
                            .iter()
                            .map(|x| match *x {
                                Object::String(..) => {
                                    decode_pdf_string(x).unwrap_or_else(String::new)
                                }
                                Object::Array(ref arr) => {
                                    decode_pdf_string(&arr[1]).unwrap_or_else(String::new)
                                }
                                _ => String::new(),
                            })
                            .filter(|x| !x.is_empty())
                            .collect(),
                        Some(_) => Vec::new(),
                        None => Vec::new(),
                    },
                    multiselect: {
                        let flags = ChoiceFlags::from_bits_truncate(get_field_flags(field));
                        flags.intersects(ChoiceFlags::MULTISELECT)
                    },
                    readonly: is_read_only(field),
                    required: is_required(field),
                }
            }
            FieldType::ComboBox => FieldState::ComboBox {
                // V field in a list box can be either text for one option, an array for many
                // options, or null
                selected: match field.get(b"V") {
                    Ok(selection) => match *selection {
                        Object::String(..) => {
                            vec![decode_pdf_string(selection).unwrap_or_else(String::new)]
                        }
                        Object::Array(ref chosen) => {
                            let mut res = Vec::new();
                            for obj in chosen {
                                if let Some(string) = decode_pdf_string(obj) {
                                    res.push(string);
                                }
                            }
                            res
                        }
                        _ => Vec::new(),
                    },
                    _ => Vec::new(),
                },
                // The options is an array of either text elements or arrays where the second
                // element is what we want
                options: match self.find_opt_in_kids_or_parents(&field) {
                    Some(&Object::Array(ref options)) => options
                        .iter()
                        .map(|x| match *x {
                            Object::String(..) => decode_pdf_string(x).unwrap_or_else(String::new),
                            Object::Array(ref arr) => {
                                if let Object::String(..) = &arr[1] {
                                    decode_pdf_string(&arr[1]).unwrap_or_else(String::new)
                                } else {
                                    String::new()
                                }
                            }
                            _ => String::new(),
                        })
                        .filter(|x| !x.is_empty())
                        .collect(),
                    Some(_) => Vec::new(),
                    None => Vec::new(),
                },
                editable: {
                    let flags = ChoiceFlags::from_bits_truncate(get_field_flags(field));

                    flags.intersects(ChoiceFlags::EDIT)
                },
                readonly: is_read_only(field),
                required: is_required(field),
            },
            FieldType::Text => FieldState::Text {
                text: match field.get(b"V") {
                    Ok(object) => decode_pdf_string(object).unwrap_or_else(String::new),
                    _ => "".to_owned(),
                },
                readonly: is_read_only(field),
                required: is_required(field),
            },
            FieldType::Unknown => FieldState::Unknown,
        }
    }

    fn find_opt_in_kids_or_parents<'a>(&'a self, field: &'a Dictionary) -> Option<&'a Object> {
        if let Some(result) = self.find_opt_in_kids(field) {
            Some(result)
        } else {
            self.find_opt_in_parents(field)
        }
    }

    fn find_opt_in_kids<'a>(&'a self, field: &'a Dictionary) -> Option<&'a Object> {
        // Look in self first
        if let Some(opt) = self.get_resolved_opt(field) {
            return Some(opt);
        }

        // Otherwise: look in /Kids
        let kids = field.get(b"Kids").ok()?.as_array().ok()?;
        for kid_ref in kids {
            if let Ok(kid_id) = kid_ref.as_reference() {
                if let Ok(kid_obj) = self.doc.get_object(kid_id) {
                    if let Ok(kid_dict) = kid_obj.as_dict() {
                        if let Some(opt) = self.get_resolved_opt(kid_dict) {
                            return Some(opt);
                        }
                    }
                }
            }
        }

        None
    }

    fn find_opt_in_parents<'a>(&'a self, field: &'a Dictionary) -> Option<&'a Object> {
        // Look in self first
        if let Some(opt) = self.get_resolved_opt(field) {
            return Some(opt);
        }

        // Finally: walk up /Parent chain
        let mut current = field;
        while let Ok(Object::Reference(parent_id)) = current.get(b"Parent") {
            if let Ok(parent_obj) = self.doc.get_object(*parent_id) {
                if let Ok(parent_dict) = parent_obj.as_dict() {
                    if let Some(opt) = self.get_resolved_opt(parent_dict) {
                        return Some(opt);
                    }
                    current = parent_dict;
                } else {
                    break;
                }
            } else {
                break;
            }
        }

        None
    }

    pub fn get_checkbox_values(&self, field: &Dictionary) -> Vec<String> {
        let mut values = Vec::new();

        if let Ok(kids) = field.get(b"Kids").and_then(|o| o.as_array()) {
            // Radio-style with multiple widgets
            for kid_ref in kids {
                if let Ok(kid_id) = kid_ref.as_reference() {
                    if let Ok(widget_obj) = self.doc.get_object(kid_id) {
                        if let Ok(widget) = widget_obj.as_dict() {
                            Self::extract_checkbox_appearances(widget, &mut values);
                        }
                    }
                }
            }
        } else {
            // Standalone checkbox (single widget)
            Self::extract_checkbox_appearances(field, &mut values);
        }

        values
    }

    fn extract_checkbox_appearances(field: &lopdf::Dictionary, values: &mut Vec<String>) {
        if let Ok(ap_dict) = field.get(b"AP").and_then(|o| o.as_dict()) {
            if let Ok(n_dict) = ap_dict.get(b"N").and_then(|o| o.as_dict()) {
                for (name, _) in n_dict.iter() {
                    if let Ok(name_str) = std::str::from_utf8(name) {
                        if name_str != "Off" && !values.contains(&name_str.to_string()) {
                            values.push(name_str.to_string());
                        }
                    }
                }
            }
        }
    }

    fn get_resolved_opt<'a>(&'a self, dict: &'a Dictionary) -> Option<&'a Object> {
        let opt_obj = dict.get(b"Opt").ok()?;

        match opt_obj {
            Object::Array(_) => Some(opt_obj),
            Object::Reference(obj_id) => self.doc.get_object(*obj_id).ok(),
            _ => None,
        }
    }

    /// If the field at index `n` is a text field, fills in that field with the text `s`.
    /// If it is not a text field, returns ValueError
    ///
    /// # Panics
    /// Will panic if n is larger than the number of fields
    pub fn set_text(&mut self, n: usize, s: String) -> Result<(), ValueError> {
        match self.get_state(n) {
            FieldState::Text { .. } => {
                let field = self
                    .doc
                    .objects
                    .get_mut(&self.form_ids[n])
                    .unwrap()
                    .as_dict_mut()
                    .unwrap();

                field.set("V", encode_pdf_string(&s));

                // Regenerate text appearance conforming the new text but ignore the result
                let _ = self.regenerate_text_appearance(n);

                Ok(())
            }
            _ => Err(ValueError::TypeMismatch),
        }
    }

    /// Regenerates the appearance for the field at index `n` due to an alteration of the
    /// original TextField value, the AP will be updated accordingly.
    ///
    /// # Incomplete
    /// This function is not exhaustive as not parse the original TextField orientation
    /// or the text alignment and other kind of enrichments, also doesn't discover for
    /// the global document DA.
    ///
    /// A more sophisticated parser is needed here
    fn regenerate_text_appearance(&mut self, n: usize) -> Result<(), lopdf::Error> {
        if let Ok(Object::Dictionary(acroform)) = self.doc.get_object_mut(self.form_ids[n]) {
            acroform.set("NeedAppearances", Object::Boolean(true));
            acroform.remove(b"AP");
        }

        let field = {
            self.doc
                .objects
                .get(&self.form_ids[n])
                .unwrap()
                .as_dict()
                .unwrap()
        };

        // The value of the object (should be a string)
        let value = field.get(b"V")?.to_owned();

        // The default appearance of the object (should be a string)
        let da = field.get(b"DA")?.to_owned();

        // The default appearance of the object (should be a string)
        let rect = field
            .get(b"Rect")?
            .as_array()?
            .iter()
            .map(|object| {
                object
                    .as_f32()
                    .unwrap_or(object.as_i64().unwrap_or(0) as f32) as f32
            })
            .collect::<Vec<_>>();

        // Gets the object stream
        let object_id = field.get(b"AP")?.as_dict()?.get(b"N")?.as_reference()?;
        let stream = self.doc.get_object_mut(object_id)?.as_stream_mut()?;

        // Decode and get the content, even if is compressed
        let mut content = {
            if let Ok(content) = stream.decompressed_content() {
                Content::decode(&content)?
            } else {
                Content::decode(&stream.content)?
            }
        };

        // Ignored operators
        let ignored_operators = vec![
            "bt", "tc", "tw", "tz", "g", "tm", "tr", "tf", "tj", "et", "q", "bmc", "emc",
        ];

        // Remove these ignored operators as we have to generate the text and fonts again
        content.operations.retain(|operation| {
            !ignored_operators.contains(&operation.operator.to_lowercase().as_str())
        });

        // Let's construct the text widget
        content.operations.append(&mut vec![
            Operation::new("BMC", vec!["Tx".into()]),
            Operation::new("q", vec![]),
            Operation::new("BT", vec![]),
        ]);

        let font = parse_font(match da {
            Object::String(ref bytes, _) => {
                Some(from_utf8(bytes).map_err(|_err| lopdf::Error::TextStringDecode)?)
            }
            _ => None,
        });

        // Define some helping font variables
        let font_name = (font.0).0;
        let font_size = (font.0).1;
        let font_color = font.1;

        // Set the font type and size and color
        content.operations.append(&mut vec![
            Operation::new("Tf", vec![font_name.into(), font_size.into()]),
            Operation::new(
                font_color.0,
                match font_color.0 {
                    "k" => vec![
                        font_color.1.into(),
                        font_color.2.into(),
                        font_color.3.into(),
                        font_color.4.into(),
                    ],
                    "rg" => vec![
                        font_color.1.into(),
                        font_color.2.into(),
                        font_color.3.into(),
                    ],
                    _ => vec![font_color.1.into()],
                },
            ),
        ]);

        // Calcolate the text offset
        let x = 2.0; // Suppose this fixed offset as we should have known the border here

        // Formula picked up from Poppler
        let dy = rect[1] - rect[3];
        let y = if dy > 0.0 {
            0.5 * dy - 0.4 * font_size as f32
        } else {
            0.5 * font_size as f32
        };

        // Set the text bounds, first are fixed at "1 0 0 1" and then the calculated x,y
        content.operations.append(&mut vec![Operation::new(
            "Tm",
            vec![1.into(), 0.into(), 0.into(), 1.into(), x.into(), y.into()],
        )]);

        // Set the text value and some finalizing operations
        content.operations.append(&mut vec![
            Operation::new("Tj", vec![value]),
            Operation::new("ET", vec![]),
            Operation::new("Q", vec![]),
            Operation::new("EMC", vec![]),
        ]);

        // Set the new content to the original stream and compress it
        if let Ok(encoded_content) = content.encode() {
            stream.set_plain_content(encoded_content);
            let _ = stream.compress();
        }

        Ok(())
    }

    /// If the field at index `n` is a checkbox field, toggles the check box based on the value
    /// `is_checked`.
    /// If it is not a checkbox field, returns ValueError
    ///
    /// # Panics
    /// Will panic if n is larger than the number of fields
    pub fn set_check_box(
        &mut self,
        n: usize,
        value: &str,
        is_checked: bool,
    ) -> Result<(), ValueError> {
        let field_id = self.form_ids[n];

        // Extract kids BEFORE mutable borrow
        let kids = self
            .doc
            .objects
            .get(&field_id)
            .and_then(|obj| obj.as_dict().ok())
            .and_then(|dict| dict.get(b"Kids").ok())
            .and_then(|o| o.as_array().ok())
            .map(|arr| arr.clone());

        if let Some(kids) = kids {
            // Now safe to borrow mutably
            for kid_ref in kids {
                if let Ok(kid_id) = kid_ref.as_reference() {
                    if let Ok(widget_obj) = self.doc.get_object_mut(kid_id) {
                        if let Ok(widget) = widget_obj.as_dict_mut() {
                            if let Ok(ap) = widget.get(b"AP").and_then(|o| o.as_dict()) {
                                if let Ok(n_dict) = ap.get(b"N").and_then(|o| o.as_dict()) {
                                    for (name, _) in n_dict.iter() {
                                        if let Ok(name_str) = std::str::from_utf8(name) {
                                            if name_str == value {
                                                let state = if is_checked {
                                                    Object::Name(name.to_vec())
                                                } else {
                                                    Object::Name(b"Off".to_vec())
                                                };
                                                widget.set("AS", state);
                                                return Ok(());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Err(ValueError::UnknownOption)
        } else {
            // Now safe to mutably access `field`
            let field = self
                .doc
                .objects
                .get_mut(&field_id)
                .unwrap()
                .as_dict_mut()
                .unwrap();

            let on = get_on_value(field);
            let state = Object::Name(if is_checked { on.as_bytes() } else { b"Off" }.to_vec());
            field.set("V", state.clone());
            field.set("AS", state);
            Ok(())
        }
    }

    /// If the field at index `n` is a radio field, toggles the radio button based on the value
    /// `choice`
    /// If it is not a radio button field or the choice is not a valid option, returns ValueError
    ///
    /// # Panics
    /// Will panic if n is larger than the number of fields
    pub fn set_radio(&mut self, n: usize, choice: String) -> Result<(), ValueError> {
        match self.get_state(n) {
            FieldState::Radio { options, .. } => {
                if options.contains(&choice) {
                    let field = self
                        .doc
                        .objects
                        .get_mut(&self.form_ids[n])
                        .unwrap()
                        .as_dict_mut()
                        .unwrap();
                    field.set("V", Object::Name(choice.into_bytes()));
                    Ok(())
                } else {
                    Err(ValueError::InvalidSelection)
                }
            }
            _ => Err(ValueError::TypeMismatch),
        }
    }

    /// If the field at index `n` is a listbox field, selects the options in `choice`
    /// If it is not a listbox field or one of the choices is not a valid option, or if too many choices are selected, returns ValueError
    ///
    /// # Panics
    /// Will panic if n is larger than the number of fields
    pub fn set_list_box(&mut self, n: usize, choices: Vec<String>) -> Result<(), ValueError> {
        match self.get_state(n) {
            FieldState::ListBox {
                options,
                multiselect,
                ..
            } => {
                if choices.iter().fold(true, |a, h| options.contains(h) && a) {
                    if !multiselect && choices.len() > 1 {
                        Err(ValueError::TooManySelected)
                    } else {
                        let field = self
                            .doc
                            .objects
                            .get_mut(&self.form_ids[n])
                            .unwrap()
                            .as_dict_mut()
                            .unwrap();
                        match choices.len() {
                            0 => field.set("V", Object::Null),
                            1 => field.set("V", encode_pdf_string(&choices[0])),
                            _ => field.set(
                                "V",
                                Object::Array(
                                    choices.iter().map(|x| encode_pdf_string(&x)).collect(),
                                ),
                            ),
                        };
                        Ok(())
                    }
                } else {
                    Err(ValueError::InvalidSelection)
                }
            }
            _ => Err(ValueError::TypeMismatch),
        }
    }

    /// If the field at index `n` is a combobox field, selects the options in `choice`
    /// If it is not a combobox field or one of the choices is not a valid option, or if too many choices are selected, returns ValueError
    ///
    /// # Panics
    /// Will panic if n is larger than the number of fields
    pub fn set_combo_box(&mut self, n: usize, choice: String) -> Result<(), ValueError> {
        match self.get_state(n) {
            FieldState::ComboBox {
                options, editable, ..
            } => {
                if options.contains(&choice) || editable {
                    let value_obj = encode_pdf_string(&choice);

                    if let Some(Object::Dictionary(field)) =
                        self.doc.objects.get_mut(&self.form_ids[n])
                    {
                        field.set("V", value_obj.clone());
                        field.set("DV", value_obj);
                        field.remove(b"AP");
                    }

                    // self.regenerate_choice_appearance(n).map_err(|_| ValueError::InvalidSelection)?;
                    Ok(())
                } else {
                    Err(ValueError::InvalidSelection)
                }
            }
            _ => Err(ValueError::TypeMismatch),
        }
    }

    /// Saves the form to the specified path
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<(), io::Error> {
        self.doc.save(path).map(|_| ())
    }

    /// Saves the form to the specified path
    pub fn save_to<W: Write>(&mut self, target: &mut W) -> Result<(), io::Error> {
        self.doc.save_to(target)
    }

    fn get_possibilities(&self, oid: ObjectId) -> Vec<String> {
        let mut res = Vec::new();
        let kids_obj = self
            .doc
            .objects
            .get(&oid)
            .unwrap()
            .as_dict()
            .unwrap()
            .get(b"Kids");
        if let Ok(&Object::Array(ref kids)) = kids_obj {
            for (i, kid) in kids.iter().enumerate() {
                let mut found = false;
                if let Ok(&Object::Dictionary(ref appearance_states)) =
                    kid.deref(&self.doc).unwrap().as_dict().unwrap().get(b"AP")
                {
                    if let Ok(&Object::Dictionary(ref normal_appearance)) =
                        appearance_states.get(b"N")
                    {
                        for (key, _) in normal_appearance {
                            if key != b"Off" {
                                res.push(from_utf8(key).unwrap_or("").to_owned());
                                found = true;
                                break;
                            }
                        }
                    }
                }

                if !found {
                    res.push(i.to_string());
                }
            }
        }

        res
    }

    pub fn update_pdf_field_value_preserving_structure(
        &mut self,
        field_id: ObjectId,
        value: Object,
        update_default: bool,
    ) {
        // First: extract any referenced object IDs from the field dict
        let (v_ref_opt, dv_ref_opt) = {
            let field_obj = self.doc.objects.get(&field_id);
            let mut v_ref = None;
            let mut dv_ref = None;

            if let Some(Object::Dictionary(field_dict)) = field_obj {
                if let Ok(&Object::Reference(r)) = field_dict.get(b"V") {
                    v_ref = Some(r);
                }
                if update_default {
                    if let Ok(&Object::Reference(r)) = field_dict.get(b"DV") {
                        dv_ref = Some(r);
                    }
                }
            }

            (v_ref, dv_ref)
        };

        let mut updated = false;

        // Then: perform mutations
        if let Some(v_ref) = v_ref_opt {
            if let Some(obj) = self.doc.objects.get_mut(&v_ref) {
                *obj = value.clone();
                updated = true;
            }
        }

        if update_default {
            if let Some(dv_ref) = dv_ref_opt {
                if let Some(obj) = self.doc.objects.get_mut(&dv_ref) {
                    *obj = value.clone();
                }
            }
        }

        if updated {
            return;
        }

        // Fallback: insert new indirect object and update field dict
        let new_value_id = self.doc.new_object_id();
        self.doc.objects.insert(new_value_id, value.clone());

        if let Some(Object::Dictionary(field_dict)) = self.doc.objects.get_mut(&field_id) {
            field_dict.set("V", Object::Reference(new_value_id));

            if update_default {
                field_dict.set("DV", Object::Reference(new_value_id));
            }

            // Optional: remove appearance stream to force redraw
            field_dict.remove(b"AP");
        }
    }

    fn regenerate_choice_appearance(&mut self, n: usize) -> Result<(), lopdf::Error> {
        let field_id = self.form_ids[n];
        let field = self
            .doc
            .objects
            .get(&field_id)
            .ok_or_else(|| lopdf::Error::DictKey("field".into()))?
            .as_dict()?;

        let value = field.get(b"V")?.clone();
        let rect = field
            .get(b"Rect")?
            .as_array()?
            .iter()
            .map(|o| o.as_f32().unwrap_or(o.as_i64().unwrap_or(0) as f32))
            .collect::<Vec<_>>();

        let da = field
            .get(b"DA")
            .or_else(|_| {
                self.doc
                    .catalog()?
                    .get(b"AcroForm")
                    .ok()
                    .and_then(|obj| obj.as_reference().ok())
                    .and_then(|id| self.doc.get_object(id).ok())
                    .and_then(|obj| obj.as_dict().ok())
                    .and_then(|acroform| acroform.get(b"DA").ok())
                    .ok_or(lopdf::Error::Unimplemented("No DA found in AcroForm"))
            })?
            .clone();

        let mut content = Content { operations: vec![] };

        content
            .operations
            .push(Operation::new("BMC", vec!["Tx".into()]));
        content.operations.push(Operation::new("q", vec![]));
        content.operations.push(Operation::new("BT", vec![]));

        // Parse font info from DA string
        let font = parse_font(match da {
            Object::String(ref bytes, _) => Some(from_utf8(bytes).unwrap_or_default()),
            _ => None,
        });
        let font_name = (font.0).0;
        let font_size = (font.0).1;
        let font_color = font.1;

        content.operations.push(Operation::new(
            "Tf",
            vec![font_name.into(), font_size.into()],
        ));
        content.operations.push(Operation::new(
            font_color.0,
            match font_color.0 {
                "k" => vec![
                    font_color.1.into(),
                    font_color.2.into(),
                    font_color.3.into(),
                    font_color.4.into(),
                ],
                "rg" => vec![
                    font_color.1.into(),
                    font_color.2.into(),
                    font_color.3.into(),
                ],
                _ => vec![font_color.1.into()],
            },
        ));

        let x = 2.0;
        let dy = rect[1] - rect[3];
        let y = if dy > 0.0 {
            0.5 * dy - 0.4 * font_size as f32
        } else {
            0.5 * font_size as f32
        };

        content.operations.push(Operation::new(
            "Tm",
            vec![1.into(), 0.into(), 0.into(), 1.into(), x.into(), y.into()],
        ));

        let string_obj = match value {
            Object::Reference(id) => match self.doc.get_object(id)? {
                Object::String(bytes, format) => Object::String(bytes.clone(), *format),
                _ => {
                    return Err(lopdf::Error::DictKey(
                        "Expected string behind /V reference".into(),
                    ));
                }
            },
            Object::String(_, _) => value.clone(),
            _ => {
                return Err(lopdf::Error::DictKey(
                    "Expected /V to be string or ref to string".into(),
                ));
            }
        };

        content
            .operations
            .push(Operation::new("Tj", vec![string_obj]));

        content.operations.push(Operation::new("ET", vec![]));
        content.operations.push(Operation::new("Q", vec![]));
        content.operations.push(Operation::new("EMC", vec![]));

        let appearance_stream = lopdf::Stream::new(
            dictionary! {
                "Subtype" => "Form",
                "Type" => "XObject",
                "BBox" => lopdf::Object::Array(rect.iter().copied().map(|f| f.into()).collect()),
                "Resources" => lopdf::dictionary! {}, // could inherit from AcroForm later
            },
            content.encode()?,
        );

        let stream_id = self.doc.add_object(appearance_stream);

        if let Some(Object::Dictionary(field_dict)) = self.doc.objects.get_mut(&field_id) {
            field_dict.set(
                "AP",
                dictionary! {
                    "N" => Object::Reference(stream_id),
                },
            );
            field_dict.set("AS", Object::Name(b"N".to_vec()));
        }

        Ok(())
    }
}
