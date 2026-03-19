use crate::error::Result;
use crate::model::{
    CtrlHeader, ListHeader, PageDef, ParaCharShape, ParaLineSeg, ParaText, Paragraph, Section,
    SectionDef, Table,
};
use crate::parser::record::{HwpTag, Record};
use crate::reader::StreamReader;
use crate::utils::compression::decompress_stream;

pub struct BodyTextParser;

impl BodyTextParser {
    pub fn parse(data: Vec<u8>, is_compressed: bool) -> Result<BodyText> {
        let data = if is_compressed {
            decompress_stream(&data)?
        } else {
            data
        };

        let mut reader = StreamReader::new(data);
        let mut sections = Vec::new();
        let mut current_section = Section::default();
        let mut current_paragraph: Option<Paragraph> = None;

        let mut first_section = true;
        // Track table context via record level
        let mut table_level: Option<u8> = None;

        while reader.remaining() >= 4 {
            // Need at least 4 bytes for record header
            let record = match Record::parse(&mut reader) {
                Ok(r) => r,
                Err(_) => break, // Stop parsing on error
            };

            current_section.debug_tags.push(record.tag_id());

            match HwpTag::from_u16(record.tag_id()) {
                // Tag 0x42 (HWPTAG_PARA_HEADER) - Paragraph header with properties
                Some(HwpTag::SectionDefine) => {
                    if first_section {
                        // First record may be section definition
                        current_section.section_def = SectionDef::from_record(&record).ok();
                        first_section = false;
                    }
                    // Check if this paragraph is inside a table cell
                    let is_cell = if let Some(tl) = table_level {
                        if record.header.level <= tl {
                            // Same or shallower level = no longer in table
                            table_level = None;
                            false
                        } else {
                            true
                        }
                    } else {
                        false
                    };
                    // Push previous paragraph and start a new one
                    if let Some(para) = current_paragraph.take() {
                        current_section.paragraphs.push(para);
                    }
                    // Parse paragraph header properties from this record
                    let mut new_para = Paragraph::from_header_record(&record)
                        .unwrap_or_default();
                    new_para.in_table = is_cell;
                    current_paragraph = Some(new_para);
                }

                // Tag 0x43 (HWPTAG_PARA_TEXT) - Paragraph text content
                Some(HwpTag::ColumnDefine) => {
                    if let Some(ref mut para) = current_paragraph {
                        if let Ok(text) = ParaText::from_record(&record) {
                            para.text = Some(text);
                        }
                    }
                }

                // Tag 0x44 (HWPTAG_PARA_CHAR_SHAPE) - Character shape positions
                Some(HwpTag::TableControl) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.char_shapes = ParaCharShape::from_record(&record).ok();
                    }
                }

                // Tag 0x45 - Skip for now (line segment parsing may produce bad data)
                Some(HwpTag::SheetControl) => {
                    // Intentionally skipped - let layout engine calculate lines dynamically
                }

                // ============================================================
                // Body text tags at CORRECT HWP 5.0 spec offsets (0x46-0x4D)
                // The HwpTag enum names don't match their body text function
                // because the same IDs have different meanings in DocInfo.
                // ============================================================

                // Tag 0x46 (HWPTAG_PARA_RANGE_TAG) - enum: none (falls to raw match below)

                // Tag 0x47 (HWPTAG_CTRL_HEADER) - enum: LineInfo
                Some(HwpTag::LineInfo) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.ctrl_header = CtrlHeader::from_record(&record).ok();
                    }
                }

                // Tag 0x48 (HWPTAG_LIST_HEADER) - enum: HiddenComment
                Some(HwpTag::HiddenComment) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.list_header = ListHeader::from_record(&record).ok();
                    }
                }

                // Tag 0x49 (HWPTAG_PAGE_DEF) - enum: HeaderFooter
                Some(HwpTag::HeaderFooter) => {
                    current_section.page_def = PageDef::from_record(&record).ok();
                }

                // Tag 0x4D (HWPTAG_TABLE) - enum: PageHide
                Some(HwpTag::PageHide) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.table_data = Table::from_record(&record).ok();
                        table_level = Some(record.header.level);
                    }
                }

                // ============================================================
                // "High" enum tag IDs (0x50+) - these may appear in some files
                // with different tag numbering schemes.
                // ============================================================

                Some(HwpTag::ParaHeader) => {
                    let is_cell = if let Some(tl) = table_level {
                        if record.header.level <= tl {
                            table_level = None;
                            false
                        } else {
                            true
                        }
                    } else {
                        false
                    };
                    if let Some(para) = current_paragraph.take() {
                        current_section.paragraphs.push(para);
                    }
                    if let Ok(mut para) = Paragraph::from_header_record(&record) {
                        para.in_table = is_cell;
                        current_paragraph = Some(para);
                    }
                }
                Some(HwpTag::ParaText) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.text = Some(ParaText::from_record(&record)?);
                    }
                }
                Some(HwpTag::ParaCharShape) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.char_shapes = ParaCharShape::from_record(&record).ok();
                    }
                }
                Some(HwpTag::ParaLineSeg) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.line_segments = ParaLineSeg::from_record(&record).ok();
                    }
                }
                Some(HwpTag::ParaRangeTag) => {
                    if let Some(ref mut para) = current_paragraph {
                        if let Ok(hyperlink) =
                            crate::model::hyperlink::Hyperlink::from_record(&record)
                        {
                            para.hyperlinks.push(hyperlink);
                        }
                    }
                }
                Some(HwpTag::CtrlHeader) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.ctrl_header = CtrlHeader::from_record(&record).ok();
                    }
                }
                Some(HwpTag::ListHeader) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.list_header = ListHeader::from_record(&record).ok();
                    }
                }
                Some(HwpTag::PageDef) => {
                    current_section.page_def = PageDef::from_record(&record).ok();
                }
                Some(HwpTag::Table) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.table_data = Table::from_record(&record).ok();
                        table_level = Some(record.header.level);
                    }
                }

                _ => {
                    // Skip other tags for now
                }
            }
        }

        // Add last paragraph and section
        if let Some(para) = current_paragraph {
            current_section.paragraphs.push(para);
        }
        // Always add the section even if empty - there's at least one section
        sections.push(current_section);

        Ok(BodyText { sections })
    }
}

#[derive(Debug, Default)]
pub struct BodyText {
    pub sections: Vec<Section>,
}

impl BodyText {
    pub fn extract_text(&self) -> String {
        let mut result = String::new();

        for section in &self.sections {
            for para in &section.paragraphs {
                if let Some(ref text) = para.text {
                    result.push_str(&text.content);
                    result.push('\n');
                }
            }
        }

        result
    }
}
