mod utils;

use crate::utils::*;
use chrono::Local;
use derive_error::Error;
use lopdf::content::{Content, Operation};
use lopdf::{Dictionary, Document, Object, ObjectId};
use lopdf::{Stream, dictionary};
use std::collections::{HashSet, VecDeque};
use std::fmt::Debug;
use std::io;
use std::io::Write;
use std::path::Path;
use std::str;

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
    CheckBoxGroup {
        states: Vec<SingleCheckboxState>,
        options: Vec<String>,
    },
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
            Ok(obj) => decode_pdf_string(obj),
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
                    let mut choices = HashSet::new();

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

                                                    let value = decode_pdf_string_from_bytes(name)
                                                        .unwrap_or_else(String::new);
                                                    choices.insert(value.clone());
                                                    checkboxes.push(SingleCheckboxState {
                                                        widget_id: kid_id,
                                                        on_value: value,
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

                    FieldState::CheckBoxGroup {
                        states: checkboxes,
                        options: Vec::from_iter(choices.iter().cloned()),
                    }
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

    fn extract_checkbox_appearances(field: &Dictionary, values: &mut Vec<String>) {
        if let Ok(ap_dict) = field.get(b"AP").and_then(|o| o.as_dict()) {
            if let Ok(n_dict) = ap_dict.get(b"N").and_then(|o| o.as_dict()) {
                for (name, _) in n_dict.iter() {
                    if let Some(name_str) = decode_pdf_string_from_bytes(name) {
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
                self.regenerate_text_appearance_safe(n)
                    .expect("Failed to set the appearance stream");

                Ok(())
            }
            _ => Err(ValueError::TypeMismatch),
        }
    }

    pub fn regenerate_text_appearance_safe(&mut self, n: usize) -> Result<(), lopdf::Error> {
        let field_id = self.form_ids[n];

        // Extract immutable data first to avoid borrow conflicts
        let (value_obj, inherited_da_str, kids_array_opt) = {
            let field_obj = self.doc.objects.get(&field_id).unwrap();
            let field_dict = field_obj.as_dict().unwrap();

            // Clone the raw value object (to pass to rendering)
            let value_obj = field_dict.get(b"V")?.clone();

            // Decode the inherited default appearance (DA) string
            let inherited_da_str = field_dict
                .get(b"DA")
                .and_then(|obj| obj.as_str())
                .map(|bytes| str::from_utf8(bytes).unwrap_or("/F1 12 Tf 0 g"))
                .unwrap_or("/F1 12 Tf 0 g");

            // Clone the /Kids array if present
            let kids_array_opt = field_dict.get(b"Kids").and_then(|o| o.as_array()).cloned();

            (value_obj, inherited_da_str.to_string(), kids_array_opt)
        };

        // Set /NeedAppearances = false to indicate manual appearance generation
        if let Ok(acroform_id) = self.doc.catalog()?.get(b"AcroForm")?.as_reference() {
            if let Ok(acroform) = self.doc.get_object_mut(acroform_id)?.as_dict_mut() {
                acroform.set(b"NeedAppearances", Object::Boolean(false));
            }
        }

        // Regenerate appearance for each widget (kid), or fallback to field itself
        if let Ok(kids_array) = kids_array_opt {
            for kid in kids_array {
                let kid_id = kid.as_reference()?;
                self.regenerate_widget_appearance(kid_id, &value_obj, &inherited_da_str)?;
            }
        } else {
            self.regenerate_widget_appearance(field_id, &value_obj, &inherited_da_str)?;
        }

        Ok(())
    }

    fn regenerate_widget_appearance(
        &mut self,
        widget_id: ObjectId,
        value: &Object,
        inherited_da: &str,
    ) -> Result<(), lopdf::Error> {
        // Decode the PDF string value into plain text, fallback to empty if invalid
        let text = decode_pdf_string(value).unwrap_or_default();

        // Extract rectangle and default appearance (DA) string
        let rect;
        let da_str;
        {
            let widget = self.doc.objects.get(&widget_id).unwrap().as_dict().unwrap();

            // Extract bounding rectangle for the widget
            rect = widget
                .get(b"Rect")?
                .as_array()?
                .iter()
                .map(|o| o.as_f32().unwrap_or(0.0))
                .collect::<Vec<_>>();

            // Try to get /DA from widget or fallback to parent or inherited
            da_str = match widget.get(b"DA") {
                Ok(da) => da,
                Err(_) => widget
                    .get(b"Parent")
                    .ok()
                    .and_then(|p| p.as_reference().ok())
                    .and_then(|pid| self.doc.objects.get(&pid))
                    .and_then(|parent| parent.as_dict().ok())
                    .and_then(|dict| dict.get(b"DA").ok())
                    .ok_or(lopdf::Error::DictKey("DA from parent".into()))?,
            }
            .as_str()
            .map(|b| str::from_utf8(b).unwrap_or(inherited_da))
            .unwrap_or(inherited_da)
            .to_string();
        }

        // Parse font name and size from DA string
        let (font, font_color) = parse_font(Some(&da_str));
        let mut font_name = font.0.to_string();
        let font_size = font.1 as f32;

        // Compute text position and BBox size relative to (0,0)
        let width = (rect[2] - rect[0]).abs();
        let height = (rect[3] - rect[1]).abs();

        // Resolve actual base font and reference from page resources
        let (resolved_name, font_ref) = self.resolve_font_from_da_name(&font_name, widget_id)?;
        font_name = String::from_utf8_lossy(&resolved_name).to_string();

        let widget = self.doc.objects.get(&widget_id).unwrap().as_dict()?;
        let max_len = get_max_len_from_widget_or_parent(&self.doc, widget);

        let font_encoding = get_font_encoding(&self.doc, font_ref);

        // Build the appearance content stream
        let mut ops = vec![
            Operation::new("q", vec![]),
            Operation::new("BT", vec![]),
            Operation::new("Tf", vec![font_name.clone().into(), font_size.into()]),
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
        ];

        // Approximate vertical offset for baseline alignment
        let bbox_width = rect[2] - rect[0];
        let bbox_height = rect[3] - rect[1];

        let ascent = font_size * 1.0;
        let line_height = font_size * 1.2;

        let padding_x = 2.0; // or bbox_width * 0.03;

        if max_len > 0 {
            // --- Cell-layout (single line, left-aligned, vertical center) ---
            let cell_width = bbox_width / max_len as f32;
            let baseline_y = (bbox_height + line_height) / 2.0 - ascent;

            for (i, ch) in text.chars().take(max_len).enumerate() {
                // instead of aligning to left of cell, offset to center
                let cell_center_x = (i as f32 + 0.5) * cell_width;
                let x = cell_center_x - font_size * 0.25; // 0.25 ≈ half of avg glyph width in em

                ops.push(Operation::new(
                    "Tm",
                    vec![
                        1.into(),
                        0.into(),
                        0.into(),
                        1.into(),
                        x.into(),
                        baseline_y.into(),
                    ],
                ));

                let encoded = encode_text_for_pdf(&ch.to_string(), font_encoding.as_deref());
                ops.push(Operation::new("Tj", vec![encoded]));
            }
        } else {
            // --- Multiline layout (left-aligned, vertical center) ---
            let lines: Vec<&str> = text.split('\n').collect();
            let block_height = lines.len() as f32 * line_height;
            let first_baseline_y = (bbox_height + block_height) / 2.0 - ascent;

            for (i, line) in lines.iter().enumerate() {
                let y = first_baseline_y - (i as f32) * line_height;
                let x = padding_x;

                ops.push(Operation::new(
                    "Tm",
                    vec![1.into(), 0.into(), 0.into(), 1.into(), x.into(), y.into()],
                ));

                let encoded = encode_text_for_pdf(line, font_encoding.as_deref());
                ops.push(Operation::new("Tj", vec![encoded]));
            }
        }

        ops.extend([Operation::new("ET", vec![]), Operation::new("Q", vec![])]);

        let content = Content { operations: ops };

        // Build font resource with the correct name and ID
        let mut font_resources = Dictionary::new();
        font_resources.set(font_name.as_bytes().to_vec(), Object::Reference(font_ref));

        let resources: Dictionary = [(b"Font".to_vec(), Object::Dictionary(font_resources))]
            .into_iter()
            .collect();

        // Build a new Form XObject stream with normalized BBox
        let stream_dict: Dictionary = [
            (b"Type".to_vec(), Object::Name(b"XObject".to_vec())),
            (b"Subtype".to_vec(), Object::Name(b"Form".to_vec())),
            (
                b"BBox".to_vec(),
                Object::Array(vec![
                    Object::Real(0.0),
                    Object::Real(0.0),
                    Object::Real(width),
                    Object::Real(height),
                ]),
            ),
            (b"Resources".to_vec(), Object::Dictionary(resources)),
        ]
        .into_iter()
        .collect();

        // Add the stream to the document
        let new_stream = Stream::new(stream_dict, content.encode()?);
        let new_stream_id = self.doc.add_object(new_stream);

        // Attach the new appearance stream to the widget annotation
        let widget = self
            .doc
            .objects
            .get_mut(&widget_id)
            .unwrap()
            .as_dict_mut()
            .unwrap();

        widget.set(b"AS", Object::Name(b"N".to_vec()));

        let timestamp = format!("D:{}", Local::now().format("%Y%m%d%H%M%S+02'00'"));
        widget.set(b"M", Object::string_literal(timestamp));

        let ap_dict = match widget.get_mut(b"AP") {
            Ok(Object::Dictionary(d)) => d,
            _ => {
                widget.set(b"AP", Object::Dictionary(Dictionary::new()));
                widget.get_mut(b"AP")?.as_dict_mut().unwrap()
            }
        };

        ap_dict.set(b"N", Object::Reference(new_stream_id));

        Ok(())
    }

    fn resolve_font_from_da_name(
        &self,
        da_font_name: &str,
        widget_id: ObjectId,
    ) -> Result<(Vec<u8>, ObjectId), lopdf::Error> {
        // Try /DR from widget or parent
        let get_dr_fonts = |dict: &Dictionary| {
            dict.get(b"DR")
                .and_then(|dr| dr.as_dict())
                .and_then(|dr| dr.get(b"Font"))
                .and_then(|f| f.as_dict())
                .map(|fonts| {
                    fonts
                        .iter()
                        .map(|(k, v)| (k.clone(), v.as_reference().unwrap()))
                        .collect::<Vec<_>>()
                })
        };

        let widget = self
            .doc
            .objects
            .get(&widget_id)
            .ok_or(lopdf::Error::Unimplemented(""))?
            .as_dict()?;

        // Try /DR from widget
        if let Ok(fonts) = get_dr_fonts(widget) {
            if let Some((k, v)) = fonts
                .iter()
                .find(|(name, _)| name == da_font_name.as_bytes())
            {
                return Ok((k.clone(), *v));
            }
        }

        // Try /Parent
        if let Ok(parent_id) = widget.get(b"Parent").and_then(Object::as_reference) {
            if let Some(parent) = self
                .doc
                .objects
                .get(&parent_id)
                .and_then(|obj| Object::as_dict(obj).ok())
            {
                if let Ok(fonts) = get_dr_fonts(parent) {
                    if let Some((k, v)) = fonts
                        .iter()
                        .find(|(name, _)| name == da_font_name.as_bytes())
                    {
                        return Ok((k.clone(), *v));
                    }
                }
            }
        }

        // Try /AcroForm DR
        let acroform_fonts = self
            .doc
            .catalog()?
            .get(b"AcroForm")
            .and_then(Object::as_reference)
            .and_then(|id| {
                self.doc
                    .objects
                    .get(&id)
                    .ok_or(lopdf::Error::Unimplemented(""))
            })
            .and_then(Object::as_dict)
            .and_then(|form| form.get(b"DR"))
            .and_then(Object::as_dict)
            .and_then(|dr| dr.get(b"Font"))
            .and_then(Object::as_dict);

        if let Ok(fonts) = acroform_fonts {
            if let Some((k, v)) = fonts
                .iter()
                .find(|(name, _)| *name == da_font_name.as_bytes())
            {
                return Ok((k.clone(), v.as_reference()?));
            }
        }

        Err(lopdf::Error::DictKey(format!(
            "Font resource '{}' not found in DR",
            da_font_name
        )))
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

        // Clone kids list
        let kids: Option<Vec<Object>> = self
            .doc
            .objects
            .get(&field_id)
            .and_then(|obj| obj.as_dict().ok())
            .and_then(|dict| dict.get(b"Kids").ok())
            .and_then(|o| o.as_array().ok())
            .map(|arr| arr.clone());

        if let Some(kids) = kids {
            let mut matched = None;

            for kid_ref in kids {
                if let Ok(kid_id) = kid_ref.as_reference() {
                    if let Ok(widget_obj) = self.doc.get_object_mut(kid_id) {
                        if let Ok(widget) = widget_obj.as_dict_mut() {
                            if let Ok(ap) = widget.get(b"AP").and_then(|o| o.as_dict()) {
                                if let Ok(n_dict) = ap.get(b"N").and_then(|o| o.as_dict()) {
                                    let mut new_state = None;

                                    for (name, _) in n_dict.iter() {
                                        if let Some(name_str) = decode_pdf_string_from_bytes(name) {
                                            if name_str == value {
                                                matched = Some(name.to_vec());
                                                new_state = Some(if is_checked {
                                                    Object::Name(name.to_vec())
                                                } else {
                                                    Object::Name(b"Off".to_vec())
                                                });
                                            }
                                        }
                                    }

                                    if let Some(state) = new_state {
                                        widget.set("AS", state);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if let Some(name) = matched {
                // optional: update /V to the currently selected value
                let field = self
                    .doc
                    .objects
                    .get_mut(&field_id)
                    .unwrap()
                    .as_dict_mut()
                    .unwrap();
                if is_checked {
                    field.set("V", Object::Name(name));
                }
                Ok(())
            } else {
                Err(ValueError::UnknownOption)
            }
        } else {
            // Single checkbox fallback
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

                    self.regenerate_choice_appearance(n)
                        .map_err(|_| ValueError::InvalidSelection)?;
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
                                res.push(
                                    decode_pdf_string_from_bytes(key).unwrap_or_else(String::new),
                                );
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

    pub fn regenerate_choice_appearance(&mut self, n: usize) -> Result<(), lopdf::Error> {
        let field_id = self.form_ids[n];
        let field = self
            .doc
            .objects
            .get(&field_id)
            .ok_or_else(|| lopdf::Error::DictKey("field".into()))?
            .as_dict()?;

        let kids = field.get(b"Kids").and_then(|o| o.as_array()).ok();

        let widget_ids = match kids {
            Some(array) => array
                .iter()
                .filter_map(|o| o.as_reference().ok())
                .collect::<Vec<_>>(),
            None => vec![field_id],
        };

        let font_ref = self
            .doc
            .objects
            .iter()
            .find_map(|(id, obj)| match obj {
                Object::Dictionary(d)
                    if d.get(b"Type")
                        .ok()
                        .map_or(false, |t| t.as_name().ok() == Some(b"Font")) =>
                {
                    Some(*id)
                }
                _ => None,
            })
            .ok_or_else(|| lopdf::Error::DictKey("Missing font resource".into()))?;

        let font_encoding = get_font_encoding(&self.doc, font_ref);

        let value_obj = field.get(b"V")?.clone();
        let value = match value_obj {
            Object::String(ref bytes, _) => decode_pdf_string_from_bytes(bytes),
            Object::Name(ref name) => decode_pdf_string_from_bytes(name),
            _ => return Err(lopdf::Error::DictKey("V".into())),
        };

        let da_string = field.get(b"DA")?.clone();
        for widget_id in widget_ids {
            let widget = self
                .doc
                .objects
                .get_mut(&widget_id)
                .ok_or_else(|| lopdf::Error::DictKey("widget".into()))?
                .as_dict_mut()?;

            widget.set("DA", da_string.clone());
            widget.set("AS", Object::Name(b"N".to_vec()));

            let rect = widget
                .get(b"Rect")?
                .as_array()?
                .iter()
                .map(|o| o.as_f32().unwrap_or(0.0))
                .collect::<Vec<_>>();

            let width = rect[2] - rect[0];
            let height = rect[3] - rect[1];

            let font_size = 11.0;

            let mut ops = vec![
                Operation::new("q", vec![]),
                Operation::new("BT", vec![]),
                Operation::new("Tf", vec!["F1".into(), font_size.into()]),
                Operation::new("rg", vec![0.into(), 0.into(), 0.into()]), // black
                Operation::new(
                    "Td",
                    vec![2.into(), ((height / 2.0) - font_size / 2.0).into()],
                ),
            ];

            // Properly encoded text object (WinAnsi or UTF-16 w/ BOM)
            let value = value.to_owned().unwrap_or_else(String::new);
            let encoded = encode_text_for_pdf(&value, font_encoding.as_deref());
            ops.push(Operation::new("Tj", vec![encoded]));

            ops.extend([Operation::new("ET", vec![]), Operation::new("Q", vec![])]);

            let content = Content { operations: ops };
            let stream_bytes = content.encode()?;

            let stream_dict = dictionary! {
                "Type" => "XObject",
                "Subtype" => "Form",
                "BBox" => vec![0.0.into(), 0.0.into(), width.into(), height.into()],
                "Resources" => dictionary! {
                    "Font" => dictionary! {
                        b"F1" => font_ref,
                    }
                },
            };

            let ap_stream = Stream::new(stream_dict, stream_bytes);

            // Drop mutable borrow of `widget` before mutably accessing `self.doc.objects`
            let widget_id_for_set = widget_id;

            // Now it's safe to insert
            let new_id = self
                .doc
                .max_id
                .checked_add(1)
                .expect("PDF object ID overflow");
            self.doc.max_id = new_id;
            self.doc
                .objects
                .insert((new_id, 0), Object::Stream(ap_stream));
            let appearance_id = Object::Reference((new_id, 0));

            // Set `AP` on widget again
            let widget = self
                .doc
                .objects
                .get_mut(&widget_id_for_set)
                .ok_or_else(|| lopdf::Error::DictKey("widget reborrow".into()))?
                .as_dict_mut()?;

            widget.set(
                "AP",
                Object::Dictionary(dictionary! {
                    "N" => appearance_id
                }),
            );
        }

        Ok(())
    }
}
