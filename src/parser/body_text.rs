use crate::error::Result;
use crate::model::{
    CtrlHeader, ListHeader, PageDef, ParaCharShape, ParaLineSeg, ParaText, Paragraph, Section,
    SectionDef, Table, TableCell,
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
        // Track table context: index of the paragraph that owns the current table
        let mut table_para_idx: Option<usize> = None;
        let mut table_cell_idx: usize = 0;
        let mut table_expected_cells: usize = 0;

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
                    // Push previous paragraph and start a new one
                    if let Some(para) = current_paragraph.take() {
                        current_section.paragraphs.push(para);
                    }
                    // Parse paragraph header properties from this record
                    let new_para = Paragraph::from_header_record(&record)
                        .unwrap_or_default();
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
                    if let Some(tpi) = table_para_idx {
                        if table_cell_idx < table_expected_cells {
                            // Parse cell from LIST_HEADER record
                            // HWP 5.0 cell LIST_HEADER: nPara(u16) + properties(u32) + cell fields
                            let mut rdr = record.data_reader();
                            if rdr.remaining() >= 32 {
                                let _para_count = rdr.read_u16().unwrap_or(0);
                                let _properties = rdr.read_u32().unwrap_or(0);
                                let col_addr = rdr.read_u16().unwrap_or(0);
                                let row_addr = rdr.read_u16().unwrap_or(0);
                                let col_span = rdr.read_u16().unwrap_or(0);
                                let row_span = rdr.read_u16().unwrap_or(0);
                                let width = rdr.read_u32().unwrap_or(0);
                                let height = rdr.read_u32().unwrap_or(0);
                                let left_margin = rdr.read_u16().unwrap_or(0);
                                let right_margin = rdr.read_u16().unwrap_or(0);
                                let top_margin = rdr.read_u16().unwrap_or(0);
                                let bottom_margin = rdr.read_u16().unwrap_or(0);
                                let border_fill_id = rdr.read_u16().unwrap_or(0);

                                let cell = TableCell {
                                    list_header_id: 0,
                                    col_span,
                                    row_span,
                                    width,
                                    height,
                                    left_margin,
                                    right_margin,
                                    top_margin,
                                    bottom_margin,
                                    border_fill_id,
                                    text_width: width.saturating_sub(
                                        (left_margin as u32) + (right_margin as u32),
                                    ),
                                    field_name: String::new(),
                                    paragraph_list_id: None,
                                    cell_address: (row_addr, col_addr),
                                };

                                if let Some(ref mut table) =
                                    current_section.paragraphs[tpi].table_data
                                {
                                    table.cells.push(cell);
                                }
                                table_cell_idx += 1;
                                if table_cell_idx >= table_expected_cells {
                                    table_para_idx = None;
                                }
                            }
                        }
                    } else if let Some(ref mut para) = current_paragraph {
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
                        if let Ok(table) = Table::from_record(&record) {
                            let expected = (table.rows as usize) * (table.cols as usize);
                            para.table_data = Some(table);
                            // Push table paragraph now so cell LIST_HEADERs can reference it
                            let para = current_paragraph.take().unwrap();
                            current_section.paragraphs.push(para);
                            table_para_idx = Some(current_section.paragraphs.len() - 1);
                            table_cell_idx = 0;
                            table_expected_cells = expected;
                        }
                    }
                }

                // ============================================================
                // "High" enum tag IDs (0x50+) - these may appear in some files
                // with different tag numbering schemes.
                // ============================================================

                Some(HwpTag::ParaHeader) => {
                    if let Some(para) = current_paragraph.take() {
                        current_section.paragraphs.push(para);
                    }
                    if let Ok(para) = Paragraph::from_header_record(&record) {
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
                    if let Some(tpi) = table_para_idx {
                        if table_cell_idx < table_expected_cells {
                            let mut rdr = record.data_reader();
                            if rdr.remaining() >= 32 {
                                let _para_count = rdr.read_u16().unwrap_or(0);
                                let _properties = rdr.read_u32().unwrap_or(0);
                                let col_addr = rdr.read_u16().unwrap_or(0);
                                let row_addr = rdr.read_u16().unwrap_or(0);
                                let col_span = rdr.read_u16().unwrap_or(0);
                                let row_span = rdr.read_u16().unwrap_or(0);
                                let width = rdr.read_u32().unwrap_or(0);
                                let height = rdr.read_u32().unwrap_or(0);
                                let left_margin = rdr.read_u16().unwrap_or(0);
                                let right_margin = rdr.read_u16().unwrap_or(0);
                                let top_margin = rdr.read_u16().unwrap_or(0);
                                let bottom_margin = rdr.read_u16().unwrap_or(0);
                                let border_fill_id = rdr.read_u16().unwrap_or(0);

                                let cell = TableCell {
                                    list_header_id: 0,
                                    col_span,
                                    row_span,
                                    width,
                                    height,
                                    left_margin,
                                    right_margin,
                                    top_margin,
                                    bottom_margin,
                                    border_fill_id,
                                    text_width: width.saturating_sub(
                                        (left_margin as u32) + (right_margin as u32),
                                    ),
                                    field_name: String::new(),
                                    paragraph_list_id: None,
                                    cell_address: (row_addr, col_addr),
                                };

                                if let Some(ref mut table) =
                                    current_section.paragraphs[tpi].table_data
                                {
                                    table.cells.push(cell);
                                }
                                table_cell_idx += 1;
                                if table_cell_idx >= table_expected_cells {
                                    table_para_idx = None;
                                }
                            }
                        }
                    } else if let Some(ref mut para) = current_paragraph {
                        para.list_header = ListHeader::from_record(&record).ok();
                    }
                }
                Some(HwpTag::PageDef) => {
                    current_section.page_def = PageDef::from_record(&record).ok();
                }
                Some(HwpTag::Table) => {
                    if let Some(ref mut para) = current_paragraph {
                        if let Ok(table) = Table::from_record(&record) {
                            let expected = (table.rows as usize) * (table.cols as usize);
                            para.table_data = Some(table);
                            let para = current_paragraph.take().unwrap();
                            current_section.paragraphs.push(para);
                            table_para_idx = Some(current_section.paragraphs.len() - 1);
                            table_cell_idx = 0;
                            table_expected_cells = expected;
                        }
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
