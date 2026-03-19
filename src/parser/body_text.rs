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

        // Table context tracking: cell counter to mark sub-paragraphs as in_table
        let mut in_table_context = false;
        let mut table_cells_remaining: usize = 0;
        let mut saw_list_header = false;

        while reader.remaining() >= 4 {
            let record = match Record::parse(&mut reader) {
                Ok(r) => r,
                Err(_) => break,
            };

            current_section.debug_tags.push(record.tag_id());

            match HwpTag::from_u16(record.tag_id()) {
                // NOTE: HwpTag enum names are from DocInfo context.
                // In body text streams the same tag IDs have different meanings.
                // 0x42=PARA_HEADER, 0x43=PARA_TEXT, 0x44=PARA_CHAR_SHAPE,
                // 0x47=CTRL_HEADER, 0x48=LIST_HEADER, 0x49=PAGE_DEF, 0x4D=TABLE

                // 0x42: PARA_HEADER
                Some(HwpTag::SectionDefine) => {
                    if first_section {
                        current_section.section_def = SectionDef::from_record(&record).ok();
                        first_section = false;
                    }
                    let is_cell = check_table_cell_state(
                        &mut in_table_context,
                        &mut table_cells_remaining,
                        &mut saw_list_header,
                    );
                    if let Some(para) = current_paragraph.take() {
                        current_section.paragraphs.push(para);
                    }
                    let mut new_para = Paragraph::from_header_record(&record).unwrap_or_default();
                    new_para.in_table = is_cell;
                    current_paragraph = Some(new_para);
                }

                // 0x43: PARA_TEXT
                Some(HwpTag::ColumnDefine) => {
                    if let Some(ref mut para) = current_paragraph {
                        if let Ok(text) = ParaText::from_record(&record) {
                            para.text = Some(text);
                        }
                    }
                }

                // 0x44: PARA_CHAR_SHAPE
                Some(HwpTag::TableControl) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.char_shapes = ParaCharShape::from_record(&record).ok();
                    }
                }

                // 0x45: PARA_LINE_SEG (skipped — layout engine calculates dynamically)
                Some(HwpTag::SheetControl) => {}

                // 0x47: CTRL_HEADER
                Some(HwpTag::LineInfo) => {
                    if let Some(ref mut para) = current_paragraph {
                        para.ctrl_header = CtrlHeader::from_record(&record).ok();
                    }
                }

                // 0x48: LIST_HEADER
                Some(HwpTag::HiddenComment) => {
                    if in_table_context {
                        saw_list_header = true;
                    }
                    if let Some(ref mut para) = current_paragraph {
                        para.list_header = ListHeader::from_record(&record).ok();
                    }
                }

                // 0x49: PAGE_DEF
                Some(HwpTag::HeaderFooter) => {
                    current_section.page_def = PageDef::from_record(&record).ok();
                }

                // 0x4D: TABLE
                Some(HwpTag::PageHide) => {
                    if let Some(ref mut para) = current_paragraph {
                        if let Ok(table) = Table::from_record(&record) {
                            start_table_context(
                                &table,
                                &mut in_table_context,
                                &mut table_cells_remaining,
                                &mut saw_list_header,
                            );
                            para.table_data = Some(table);
                        }
                    }
                }

                // 0x50+ "high" tag IDs — same logic, proper enum names
                Some(HwpTag::ParaHeader) => {
                    let is_cell = check_table_cell_state(
                        &mut in_table_context,
                        &mut table_cells_remaining,
                        &mut saw_list_header,
                    );
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
                    if in_table_context {
                        saw_list_header = true;
                    }
                    if let Some(ref mut para) = current_paragraph {
                        para.list_header = ListHeader::from_record(&record).ok();
                    }
                }
                Some(HwpTag::PageDef) => {
                    current_section.page_def = PageDef::from_record(&record).ok();
                }
                Some(HwpTag::Table) => {
                    if let Some(ref mut para) = current_paragraph {
                        if let Ok(table) = Table::from_record(&record) {
                            start_table_context(
                                &table,
                                &mut in_table_context,
                                &mut table_cells_remaining,
                                &mut saw_list_header,
                            );
                            para.table_data = Some(table);
                        }
                    }
                }

                _ => {}
            }
        }

        if let Some(para) = current_paragraph {
            current_section.paragraphs.push(para);
        }
        sections.push(current_section);

        Ok(BodyText { sections })
    }
}

/// Determine if the current PARA_HEADER is inside a table cell.
/// Returns true if it's a cell sub-paragraph.
fn check_table_cell_state(
    in_table_context: &mut bool,
    table_cells_remaining: &mut usize,
    saw_list_header: &mut bool,
) -> bool {
    if *in_table_context && *saw_list_header {
        *saw_list_header = false;
        *table_cells_remaining = table_cells_remaining.saturating_sub(1);
        if *table_cells_remaining == 0 {
            *in_table_context = false;
        }
        true
    } else {
        *in_table_context
    }
}

/// Enter table context when a TABLE record is encountered.
fn start_table_context(
    table: &Table,
    in_table_context: &mut bool,
    table_cells_remaining: &mut usize,
    saw_list_header: &mut bool,
) {
    *in_table_context = true;
    *table_cells_remaining = (table.rows as usize) * (table.cols as usize);
    *saw_list_header = false;
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
