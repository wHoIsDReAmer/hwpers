use crate::error::Result;
use crate::parser::record::Record;

#[derive(Debug, Default)]
pub struct Section {
    pub paragraphs: Vec<Paragraph>,
    pub section_def: Option<crate::model::SectionDef>,
    pub page_def: Option<crate::model::PageDef>,
    /// Debug: all raw tag IDs seen during parsing
    pub debug_tags: Vec<u16>,
}

#[derive(Debug, Default)]
pub struct Paragraph {
    pub text: Option<ParaText>,
    pub control_mask: u32,
    pub para_shape_id: u16,
    pub style_id: u8,
    pub column_type: u8,
    pub char_shape_count: u16,
    pub range_tag_count: u16,
    pub line_align_count: u16,
    pub instance_id: u32,
    pub char_shapes: Option<crate::model::ParaCharShape>,
    pub line_segments: Option<crate::model::ParaLineSeg>,
    pub list_header: Option<crate::model::ListHeader>,
    pub ctrl_header: Option<crate::model::CtrlHeader>,
    // Store actual control data
    pub table_data: Option<crate::model::control::Table>,
    pub picture_data: Option<crate::model::control::Picture>,
    pub text_box_data: Option<crate::model::text_box::TextBox>,
    // Store hyperlinks for this paragraph
    pub hyperlinks: Vec<crate::model::hyperlink::Hyperlink>,
    /// True if this paragraph is inside a table cell (should not be rendered standalone)
    pub in_table: bool,
}

impl Paragraph {
    pub fn from_header_record(record: &Record) -> Result<Self> {
        let mut reader = record.data_reader();

        // HWP 5.0 HWPTAG_PARA_HEADER format:
        // u32: nChars (text character count)
        // u32: controlMask
        // u16: paraShapeId
        // u8:  styleId
        // u8:  divideType (column type)
        // u16: nCharShapeRef
        // u16: nRangeTag
        // u16: nLineAligns
        // u32: instanceId
        // Total: 22 bytes minimum

        if reader.remaining() < 22 {
            // Not enough data for full paragraph header, return default
            return Ok(Self::default());
        }

        let _n_chars = reader.read_u32()?; // text character count (skip)

        Ok(Self {
            control_mask: reader.read_u32()?,
            para_shape_id: reader.read_u16()?,
            style_id: reader.read_u8()?,
            column_type: reader.read_u8()?,
            char_shape_count: reader.read_u16()?,
            range_tag_count: reader.read_u16()?,
            line_align_count: reader.read_u16()?,
            instance_id: reader.read_u32()?,
            hyperlinks: Vec::new(),
            ..Default::default()
        })
    }

    pub fn parse_char_shapes(&mut self, _record: &Record) -> Result<()> {
        // Character shape parsing logic would go here
        // For now, we'll skip the implementation
        Ok(())
    }
}

#[derive(Debug)]
pub struct ParaText {
    pub content: String,
}

impl ParaText {
    pub fn from_record(record: &Record) -> Result<Self> {
        // Check if this is a table marker record
        if record.tag_id() == 0x43 && record.data.len() == 18 {
            // Check for the specific table marker pattern
            if record.data[0] == 0x0B
                && record.data[1] == 0x00
                && record.data[2] == 0x20
                && record.data[3] == 0x6C
                && record.data[4] == 0x62
                && record.data[5] == 0x74
            {
                // This is a table marker, return empty text
                return Ok(Self {
                    content: String::new(),
                });
            }
        }

        let mut reader = record.data_reader();
        let mut content = String::new();
        let mut chars = Vec::new();

        // Read all UTF-16LE characters
        while reader.remaining() >= 2 {
            let ch = reader.read_u16()?;
            chars.push(ch);
        }

        // Process characters based on record type
        if record.tag_id() == 0x43 {
            // For tag 0x43, we need special handling
            let mut i = 0;
            while i < chars.len() {
                let ch = chars[i];

                // Process characters
                match ch {
                    0x0000 => {
                        // Skip null characters
                    }
                    0x0001..=0x0008 | 0x000B | 0x000C => {
                        // HWP extended control characters: 8 UTF-16 words total
                        // (1 control char + 7 parameter words). Skip the 7 params.
                        i += 7;
                    }
                    0x0009 => {
                        // Tab character - check if this is followed by form field markers
                        if i + 2 < chars.len()
                            && chars[i + 2] == 0x0000
                            && (chars[i + 1] == 0x0480 || chars[i + 1] == 0x0264)
                        {
                            // ɤ followed by null
                            // This is a form field marker, skip the entire sequence
                            // Skip until we find normal text again (not tab, space, or control chars)
                            while i < chars.len()
                                && (chars[i] == 0x0009
                                    || chars[i] == 0x0020
                                    || chars[i] == 0x0480
                                    || chars[i] == 0x0100
                                    || chars[i] == 0x0264
                                    || chars[i] == 0x0000
                                    || chars[i] == 0x0001)
                            {
                                i += 1;
                            }
                            i -= 1; // Adjust because loop will increment
                            continue;
                        } else {
                            content.push('\t'); // Regular tab
                        }
                    }
                    0x000A => content.push('\n'), // Line feed
                    0x000D => content.push('\r'), // Carriage return
                    0x000E..=0x001F => {
                        // Skip other control characters
                    }
                    0x0264 => {
                        // ɤ character - check if part of form field
                        if i + 1 < chars.len() && chars[i + 1] == 0x0100 {
                            // Skip form field marker
                            i += 1; // Skip the Ā
                            continue;
                        } else {
                            // Regular character
                            if let Some(unicode_char) = std::char::from_u32(ch as u32) {
                                content.push(unicode_char);
                            }
                        }
                    }
                    0x0480 => {
                        // Ҁ character - check if part of form field
                        if i + 1 < chars.len() && chars[i + 1] == 0x0100 {
                            // Skip form field marker
                            i += 1; // Skip the Ā
                            continue;
                        } else {
                            // Regular character
                            if let Some(unicode_char) = std::char::from_u32(ch as u32) {
                                content.push(unicode_char);
                            }
                        }
                    }
                    0xF020..=0xF07F => {
                        // Extended control characters - skip
                    }
                    _ => {
                        // Regular characters
                        if let Some(unicode_char) = std::char::from_u32(ch as u32) {
                            content.push(unicode_char);
                        }
                    }
                }
                i += 1;
            }
        } else {
            // Standard text processing for other tags
            let mut i = 0;
            while i < chars.len() {
                let ch = chars[i];
                match ch {
                    0x0000 => {}
                    0x0001..=0x0008 | 0x000B | 0x000C => {
                        // HWP extended controls: skip 7 parameter words
                        i += 7;
                    }
                    0x0009 => content.push('\t'),
                    0x000A => content.push('\n'),
                    0x000D => content.push('\r'),
                    0x000E..=0x001F => {} // Single-char controls, skip
                    0xF020..=0xF07F => {} // Extended control characters
                    _ => {
                        if let Some(unicode_char) = std::char::from_u32(ch as u32) {
                            content.push(unicode_char);
                        }
                    }
                }
                i += 1;
            }
        }

        Ok(Self { content })
    }
}
