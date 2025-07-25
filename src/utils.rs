use bitflags::bitflags;
use lopdf::{Dictionary, Object, ObjectId, StringFormat, decode_text_string};

bitflags! {
    pub struct FieldFlags: u32 {
        const READONLY          = 0x1;
        const REQUIRED          = 0x2;
    }
}

bitflags! {
    pub struct ButtonFlags: u32 {
        const NO_TOGGLE_TO_OFF  = 0x8000;
        const RADIO             = 0x10000;
        const PUSHBUTTON        = 0x20000;
        const RADIO_IN_UNISON   = 0x4000000;

    }
}

bitflags! {
    pub struct ChoiceFlags: u32 {
        const COBMO             = 0x20000;
        const EDIT              = 0x40000;
        const SORT              = 0x80000;
        const MULTISELECT       = 0x200000;
        const DO_NOT_SPELLCHECK = 0x800000;
        const COMMIT_ON_CHANGE  = 0x8000000;
    }
}

pub fn is_read_only(field: &Dictionary) -> bool {
    let flags = FieldFlags::from_bits_truncate(get_field_flags(field));

    flags.intersects(FieldFlags::READONLY)
}

pub fn is_required(field: &Dictionary) -> bool {
    let flags = FieldFlags::from_bits_truncate(get_field_flags(field));

    flags.intersects(FieldFlags::REQUIRED)
}

pub fn get_field_flags(field: &Dictionary) -> u32 {
    field
        .get(b"Ff")
        .unwrap_or(&Object::Integer(0))
        .as_i64()
        .unwrap() as u32
}

pub fn get_on_value(field: &Dictionary) -> String {
    let mut option = None;
    if let Ok(ap) = field.get(b"AP") {
        if let Ok(dict) = ap.as_dict() {
            if let Ok(values) = dict.get(b"N") {
                if let Ok(options) = values.as_dict() {
                    for (name, _) in options {
                        if let Some(name) = decode_pdf_string_from_bytes(name) {
                            // TODO: Fix this
                            if name != "Off" && option.is_none() {
                                option = Some(name);
                            }
                        }
                    }
                }
            }
        }
    }

    option.unwrap_or("Yes".into())
}

pub fn parse_font(font_string: Option<&str>) -> ((&str, i32), (&str, i32, i32, i32, i32)) {
    // The default font object (/Helv 12 Tf 0 g)
    let default_font = ("Helv", 12);
    let default_color = ("g", 0, 0, 0, 0);

    // Build the font basing on the default appearance, if exists, if not,
    // assume a default font (surely to be improved!)
    match font_string {
        Some(font_string) => {
            let font = font_string
                .trim_start_matches('/')
                .split("Tf")
                .collect::<Vec<_>>();

            if font.len() < 2 {
                (default_font, default_color)
            } else {
                let font_family = font[0].trim().split(' ').collect::<Vec<_>>();
                let font_color = font[1].trim().split(' ').collect::<Vec<_>>();

                let font = if font_family.len() >= 2 {
                    (font_family[0], font_family[1].parse::<i32>().unwrap_or(0))
                } else {
                    default_font
                };

                let color = if font_color.len() == 2 {
                    ("g", font_color[0].parse::<i32>().unwrap_or(0), 0, 0, 0)
                } else if font_color.len() == 4 {
                    (
                        "rg",
                        font_color[0].parse::<i32>().unwrap_or(0),
                        font_color[1].parse::<i32>().unwrap_or(0),
                        font_color[2].parse::<i32>().unwrap_or(0),
                        0,
                    )
                } else if font_color.len() == 5 {
                    (
                        "k",
                        font_color[0].parse::<i32>().unwrap_or(0),
                        font_color[1].parse::<i32>().unwrap_or(0),
                        font_color[2].parse::<i32>().unwrap_or(0),
                        font_color[3].parse::<i32>().unwrap_or(0),
                    )
                } else {
                    default_color
                };

                (font, color)
            }
        }
        _ => (default_font, default_color),
    }
}

pub fn decode_pdf_string(obj: &Object) -> Option<String> {
    decode_text_string(obj).ok()
}

// HACK: Convert bytes into object, then decode object.
pub fn decode_pdf_string_from_bytes(bytes: &[u8]) -> Option<String> {
    let s = bytes.to_vec();
    if s.starts_with(b"\xFE\xFF") {
        let obj = Object::String(s, StringFormat::Hexadecimal);
        decode_pdf_string(&obj)
    } else if s.starts_with(b"\xEF\xBB\xBF") {
        String::from_utf8(s).ok()
    } else {
        // If neither BOM is detected, PDFDocEncoding is used
        let obj = Object::String(s, StringFormat::Literal);
        decode_pdf_string(&obj)
    }
}

pub fn encode_pdf_string(value: &str) -> Object {
    encode_text_for_pdf(value, None)
}

pub fn get_font_encoding(doc: &lopdf::Document, font_ref: ObjectId) -> Option<String> {
    let font_dict = doc.get_object(font_ref).ok()?.as_dict().ok()?;
    match font_dict.get(b"Encoding") {
        Ok(Object::Name(name)) => Some(String::from_utf8_lossy(name).into()),
        Ok(Object::Reference(id)) => doc
            .get_object(*id)
            .ok()
            .and_then(|obj| obj.as_name().ok())
            .map(|n| String::from_utf8_lossy(n).into()),
        _ => None,
    }
}

pub fn encode_text_for_pdf(text: &str, encoding: Option<&str>) -> Object {
    match encoding {
        Some("Identity-H") | None => {
            // UTF-16BE with BOM
            let utf16: Vec<u8> = text.encode_utf16().flat_map(|c| c.to_be_bytes()).collect();
            let mut with_bom = vec![0xFE, 0xFF];
            with_bom.extend(utf16);
            Object::String(with_bom, StringFormat::Hexadecimal)
        }
        Some("WinAnsiEncoding") | Some("StandardEncoding") | Some("MacRomanEncoding") | _ => {
            // Best-effort Latin-1 fallback
            let bytes: Vec<u8> = text
                .chars()
                .map(|c| c as u32)
                .map(|c| if c <= 0xFF { c as u8 } else { b'?' }) // Replace non-Latin1 chars
                .collect();
            Object::String(bytes, StringFormat::Literal)
        }
    }
}

pub fn get_max_len_from_widget_or_parent(doc: &lopdf::Document, widget: &Dictionary) -> usize {
    // Try parent first
    let parent_maxlen = widget
        .get(b"Parent")
        .ok()
        .and_then(|p| p.as_reference().ok())
        .and_then(|pid| doc.objects.get(&pid))
        .and_then(|parent| parent.as_dict().ok())
        .and_then(|dict| dict.get(b"MaxLen").ok())
        .and_then(|o| o.as_i64().ok());

    // Fallback to field itself
    parent_maxlen
        .or_else(|| widget.get(b"MaxLen").ok().and_then(|o| o.as_i64().ok()))
        .unwrap_or(0) as usize
}
