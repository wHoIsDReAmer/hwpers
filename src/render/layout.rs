use crate::model::{CharShape, ParaLineSeg, ParaShape, Paragraph, Section};
use crate::HwpDocument;

#[derive(Debug, Clone)]
pub struct RenderedLine {
    pub y: i32,
    pub height: i32,
    pub baseline_y: i32,
    pub runs: Vec<TextRun>,
}

#[derive(Debug, Clone)]
pub struct TextRun {
    pub x: i32,
    pub width: i32,
    pub text: String,
    pub char_shape_id: u16,
    pub font_size: i32,
}

#[derive(Debug, Clone)]
pub struct RenderedParagraph {
    pub x: i32,
    pub y: i32,
    pub width: i32,
    pub height: i32,
    pub lines: Vec<RenderedLine>,
    pub para_shape_id: u16,
}

#[derive(Debug, Clone)]
pub struct RenderedPage {
    pub width: u32,
    pub height: u32,
    pub paragraphs: Vec<RenderedParagraph>,
    pub page_number: u32,
}

#[derive(Debug)]
pub struct LayoutResult {
    pub pages: Vec<RenderedPage>,
    pub total_height: i32,
}

pub struct LayoutEngine<'a> {
    document: &'a HwpDocument,
}

impl<'a> LayoutEngine<'a> {
    pub fn new(document: &'a HwpDocument) -> Self {
        Self { document }
    }

    pub fn calculate_layout(&self) -> LayoutResult {
        let mut pages = Vec::new();
        let mut current_page_num = 1;

        for section in self.document.sections() {
            let section_pages = self.layout_section(section, &mut current_page_num);
            pages.extend(section_pages);
        }

        let total_height = pages.iter().map(|p| p.height as i32).sum();
        LayoutResult {
            pages,
            total_height,
        }
    }

    fn layout_section(&self, section: &Section, page_num: &mut u32) -> Vec<RenderedPage> {
        let mut pages = Vec::new();

        let default_page_def = crate::model::PageDef::new_default();
        let page_def = section.page_def.as_ref().unwrap_or(&default_page_def);

        let mut current_page = RenderedPage {
            width: page_def.width,
            height: page_def.height,
            paragraphs: Vec::new(),
            page_number: *page_num,
        };

        let content_x = page_def.left_margin as i32;
        let content_y = page_def.top_margin as i32;
        let content_width = page_def.effective_width() as i32;
        let content_height = page_def.effective_height() as i32;
        let mut current_y = content_y;

        for paragraph in &section.paragraphs {
            if paragraph.in_table {
                continue;
            }

            if let Some(rendered_para) =
                self.layout_paragraph(paragraph, content_x, current_y, content_width)
            {
                if current_y + rendered_para.height > content_y + content_height {
                    pages.push(current_page);
                    *page_num += 1;

                    current_page = RenderedPage {
                        width: page_def.width,
                        height: page_def.height,
                        paragraphs: Vec::new(),
                        page_number: *page_num,
                    };
                    current_y = content_y;
                }

                current_y += rendered_para.height;
                current_page.paragraphs.push(rendered_para);
            }
        }

        if !current_page.paragraphs.is_empty() {
            pages.push(current_page);
            *page_num += 1;
        }

        pages
    }

    fn layout_paragraph(
        &self,
        paragraph: &Paragraph,
        x: i32,
        y: i32,
        width: i32,
    ) -> Option<RenderedParagraph> {
        let text = paragraph
            .text
            .as_ref()
            .map(|t| t.content.as_str())
            .unwrap_or("");

        // Table paragraphs: estimate height from row count
        if let Some(table) = &paragraph.table_data {
            let row_height = 400i32;
            let estimated_height = (table.rows as i32).max(1) * row_height;
            return Some(RenderedParagraph {
                x,
                y,
                width,
                height: estimated_height,
                lines: Vec::new(),
                para_shape_id: paragraph.para_shape_id,
            });
        }

        if text.is_empty() {
            return None;
        }

        let default_para_shape = crate::model::ParaShape::new_default();
        let para_shape = self
            .document
            .get_para_shape(paragraph.para_shape_id as usize)
            .unwrap_or(&default_para_shape);

        let para_x = x + para_shape.left_margin;
        let para_width = width - para_shape.left_margin - para_shape.right_margin;

        if let Some(line_segs) = &paragraph.line_segments {
            let lines = self.render_with_line_segments(text, paragraph, line_segs, para_x, y);
            let total_height = line_segs.total_height();

            return Some(RenderedParagraph {
                x: para_x,
                y,
                width: para_width,
                height: total_height,
                lines,
                para_shape_id: paragraph.para_shape_id,
            });
        }

        let lines = self.calculate_line_breaks(text, paragraph, para_shape, para_x, y, para_width);
        let total_height = lines.iter().map(|l| l.height).sum();

        Some(RenderedParagraph {
            x: para_x,
            y,
            width: para_width,
            height: total_height,
            lines,
            para_shape_id: paragraph.para_shape_id,
        })
    }

    fn render_with_line_segments(
        &self,
        text: &str,
        paragraph: &Paragraph,
        line_segs: &ParaLineSeg,
        x: i32,
        y: i32,
    ) -> Vec<RenderedLine> {
        let mut lines = Vec::new();
        let mut current_y = y;

        for (idx, line_seg) in line_segs.line_segments.iter().enumerate() {
            let start_char = if idx == 0 {
                0
            } else {
                line_segs.line_segments[idx - 1].text_start_position as usize
            };
            let end_char = line_seg.text_start_position as usize;

            let total_chars = text.chars().count();
            if end_char > total_chars {
                break;
            }

            let start_byte = Self::char_index_to_byte_index(text, start_char);
            let end_byte = Self::char_index_to_byte_index(text, end_char);
            let line_text = &text[start_byte..end_byte];
            let runs = self.create_text_runs_for_line(line_text, paragraph, x, start_char as u32);

            lines.push(RenderedLine {
                y: current_y + line_seg.line_vertical_position,
                height: line_seg.line_height,
                baseline_y: current_y
                    + line_seg.line_vertical_position
                    + line_seg.distance_baseline_to_line_vertical_position,
                runs,
            });

            current_y += line_seg.line_height;
        }

        lines
    }

    fn calculate_line_breaks(
        &self,
        text: &str,
        paragraph: &Paragraph,
        para_shape: &ParaShape,
        x: i32,
        y: i32,
        _width: i32,
    ) -> Vec<RenderedLine> {
        let mut lines = Vec::new();
        let mut current_y = y + para_shape.top_para_space;

        let default_char_shape_id = paragraph
            .char_shapes
            .as_ref()
            .and_then(|cs| cs.char_positions.first())
            .map(|cs| cs.char_shape_id)
            .unwrap_or(0);

        let char_shape = self.document.get_char_shape(default_char_shape_id as usize);
        let line_height = self.calculate_line_height(para_shape, char_shape);

        for line_text in text.lines() {
            if line_text.is_empty() {
                current_y += line_height;
                continue;
            }

            let runs = self.create_text_runs_for_line(line_text, paragraph, x, 0);

            lines.push(RenderedLine {
                y: current_y,
                height: line_height,
                baseline_y: current_y + (line_height * 4 / 5),
                runs,
            });

            current_y += line_height;
        }

        lines
    }

    fn char_index_to_byte_index(s: &str, char_idx: usize) -> usize {
        s.char_indices()
            .nth(char_idx)
            .map(|(byte_pos, _)| byte_pos)
            .unwrap_or(s.len())
    }

    fn create_text_runs_for_line(
        &self,
        line_text: &str,
        paragraph: &Paragraph,
        x: i32,
        text_offset: u32,
    ) -> Vec<TextRun> {
        let mut runs = Vec::new();

        if let Some(char_shapes) = &paragraph.char_shapes {
            let mut current_x = x;
            let text_offset = text_offset as usize;
            let total_chars = line_text.chars().count();

            for (idx, pos_shape) in char_shapes.char_positions.iter().enumerate() {
                let start_char = pos_shape.position as usize;
                let end_char = if idx + 1 < char_shapes.char_positions.len() {
                    char_shapes.char_positions[idx + 1].position as usize
                } else {
                    total_chars + text_offset
                };

                if start_char >= text_offset && (start_char - text_offset) < total_chars {
                    let run_start_char = (start_char - text_offset).min(total_chars);
                    let run_end_char = (end_char - text_offset).min(total_chars);

                    if run_start_char < run_end_char {
                        let run_start_byte =
                            Self::char_index_to_byte_index(line_text, run_start_char);
                        let run_end_byte = Self::char_index_to_byte_index(line_text, run_end_char);

                        let run_text = &line_text[run_start_byte..run_end_byte];
                        let char_shape = self
                            .document
                            .get_char_shape(pos_shape.char_shape_id as usize);

                        let width = self.calculate_text_width(run_text, char_shape);

                        runs.push(TextRun {
                            x: current_x,
                            width,
                            text: run_text.to_string(),
                            char_shape_id: pos_shape.char_shape_id,
                            font_size: char_shape.map(|cs| cs.base_size).unwrap_or(1000),
                        });

                        current_x += width;
                    }
                }
            }
        } else {
            let width = self.calculate_text_width(line_text, None);
            runs.push(TextRun {
                x,
                width,
                text: line_text.to_string(),
                char_shape_id: 0,
                font_size: 1000,
            });
        }

        runs
    }

    fn calculate_line_height(&self, para_shape: &ParaShape, char_shape: Option<&CharShape>) -> i32 {
        let base_size = char_shape.map(|cs| cs.base_size).unwrap_or(1000);

        match para_shape.line_space_type {
            0 => base_size * para_shape.line_space / 100, // Percentage
            1 => para_shape.line_space,                   // Fixed
            2 => base_size.max(para_shape.line_space),    // At least
            _ => base_size,
        }
    }

    fn calculate_text_width(&self, text: &str, char_shape: Option<&CharShape>) -> i32 {
        let font_size = char_shape.map(|cs| cs.base_size).unwrap_or(1000);

        text.chars()
            .map(|ch| {
                if is_fullwidth(ch) {
                    font_size
                } else {
                    font_size / 2
                }
            })
            .sum()
    }
}

fn is_fullwidth(ch: char) -> bool {
    let cp = ch as u32;
    (0x4E00..=0x9FFF).contains(&cp)
        || (0xAC00..=0xD7AF).contains(&cp)
        || (0x1100..=0x11FF).contains(&cp)
        || (0x3130..=0x318F).contains(&cp)
        || (0xF900..=0xFAFF).contains(&cp)
        || (0x3400..=0x4DBF).contains(&cp)
        || (0xFF01..=0xFF60).contains(&cp)
        || (0x3000..=0x303F).contains(&cp)
        || (0x3040..=0x30FF).contains(&cp)
        || (0x20000..=0x2FA1F).contains(&cp)
}
