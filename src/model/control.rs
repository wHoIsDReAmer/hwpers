#[derive(Debug, Clone)]
pub enum Control {
    SectionDef,
    ColumnDef,
    Table,
    ShapeObject,
    Equation,
    Picture,
    Header,
    Footer,
    Footnote,
    Endnote,
    AutoNumber,
    NewNumber,
    PageHide,
    PageOddEvenAdjust,
    PageNumberPosition,
    IndexMark,
    BookMark,
    OverlappingLetter,
    DutmalSaero,
    HiddenComment,
}
/// Picture/Image control structure
#[derive(Debug, Clone)]
pub struct Picture {
    pub properties: u32,
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
    pub z_order: i32,
    pub outer_margin_left: u16,
    pub outer_margin_right: u16,
    pub outer_margin_top: u16,
    pub outer_margin_bottom: u16,
    pub instance_id: u32,
    pub bin_item_id: u16,
    pub border_fill_id: u16,
    pub image_width: u32,
    pub image_height: u32,
}

impl Picture {
    pub fn new_default(bin_item_id: u16, width: u32, height: u32) -> Self {
        Self {
            properties: 0x80000000, // Default picture properties
            left: 567,              // 2mm left position
            top: 567,               // 2mm top position
            right: (width as i32) + 567,
            bottom: (height as i32) + 567,
            z_order: 1,               // Above text layer
            outer_margin_left: 283,   // 1mm outer margin
            outer_margin_right: 283,  // 1mm outer margin
            outer_margin_top: 283,    // 1mm outer margin
            outer_margin_bottom: 283, // 1mm outer margin
            instance_id: bin_item_id as u32,
            bin_item_id,
            border_fill_id: 0,
            image_width: width,
            image_height: height,
        }
    }

    /// Serialize picture to bytes for HWP format
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut data = Vec::new();

        // Basic picture properties
        data.extend_from_slice(&self.properties.to_le_bytes());
        data.extend_from_slice(&self.left.to_le_bytes());
        data.extend_from_slice(&self.top.to_le_bytes());
        data.extend_from_slice(&self.right.to_le_bytes());
        data.extend_from_slice(&self.bottom.to_le_bytes());
        data.extend_from_slice(&self.z_order.to_le_bytes());
        data.extend_from_slice(&self.outer_margin_left.to_le_bytes());
        data.extend_from_slice(&self.outer_margin_right.to_le_bytes());
        data.extend_from_slice(&self.outer_margin_top.to_le_bytes());
        data.extend_from_slice(&self.outer_margin_bottom.to_le_bytes());
        data.extend_from_slice(&self.instance_id.to_le_bytes());
        data.extend_from_slice(&self.bin_item_id.to_le_bytes());
        data.extend_from_slice(&self.border_fill_id.to_le_bytes());
        data.extend_from_slice(&self.image_width.to_le_bytes());
        data.extend_from_slice(&self.image_height.to_le_bytes());

        data
    }
}

#[derive(Debug, Clone)]
pub struct Table {
    pub properties: u32,
    pub rows: u16,
    pub cols: u16,
    pub cell_spacing: u16,
    pub left_margin: i32,
    pub right_margin: i32,
    pub top_margin: i32,
    pub bottom_margin: i32,
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone)]
pub struct TableCell {
    pub list_header_id: u32,
    pub col_span: u16,
    pub row_span: u16,
    pub width: u32,
    pub height: u32,
    pub left_margin: u16,
    pub right_margin: u16,
    pub top_margin: u16,
    pub bottom_margin: u16,
    pub border_fill_id: u16,
    pub text_width: u32,
    pub field_name: String,
    /// Reference to the paragraphs that form this cell's content
    pub paragraph_list_id: Option<u32>,
    /// Cell address for easier reference (row, col)
    pub cell_address: (u16, u16),
}

impl Table {
    pub fn new_default(rows: u16, cols: u16) -> Self {
        Self {
            properties: 0x0001, // Enable border
            rows,
            cols,
            cell_spacing: 142,  // 0.5mm cell spacing
            left_margin: 567,   // 2mm left margin
            right_margin: 567,  // 2mm right margin
            top_margin: 567,    // 2mm top margin
            bottom_margin: 567, // 2mm bottom margin
            cells: Vec::new(),
        }
    }

    /// Get cell at specific row and column
    pub fn get_cell(&self, row: u16, col: u16) -> Option<&TableCell> {
        self.cells
            .iter()
            .find(|cell| cell.cell_address == (row, col))
    }

    /// Get mutable cell at specific row and column
    pub fn get_cell_mut(&mut self, row: u16, col: u16) -> Option<&mut TableCell> {
        self.cells
            .iter_mut()
            .find(|cell| cell.cell_address == (row, col))
    }

    /// Add a cell to the table
    pub fn add_cell(&mut self, row: u16, col: u16, cell: TableCell) {
        // Remove any existing cell at this position
        self.cells.retain(|c| c.cell_address != (row, col));
        self.cells.push(cell);
    }

    /// Create a basic cell at the specified position
    pub fn create_cell(&mut self, row: u16, col: u16, width: u32, height: u32) -> &mut TableCell {
        let cell = TableCell {
            list_header_id: 0,
            col_span: 1,
            row_span: 1,
            width,
            height,
            left_margin: 100,
            right_margin: 100,
            top_margin: 100,
            bottom_margin: 100,
            border_fill_id: 0,
            text_width: width.saturating_sub(200), // width - margins
            field_name: format!("cell_{}_{}", row, col),
            paragraph_list_id: None,
            cell_address: (row, col),
        };

        self.add_cell(row, col, cell);
        self.get_cell_mut(row, col).unwrap()
    }

    /// Set paragraph list ID for a cell (links cell to its content paragraphs)
    pub fn set_cell_paragraph_list(&mut self, row: u16, col: u16, paragraph_list_id: u32) -> bool {
        if let Some(cell) = self.get_cell_mut(row, col) {
            cell.paragraph_list_id = Some(paragraph_list_id);
            true
        } else {
            false
        }
    }

    /// Get all cells in row order
    pub fn cells_by_row(&self) -> Vec<&TableCell> {
        let mut cells = self.cells.iter().collect::<Vec<_>>();
        cells.sort_by_key(|cell| (cell.cell_address.0, cell.cell_address.1));
        cells
    }

    /// Serialize table to bytes for HWP format
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut data = Vec::new();

        // Basic table properties (24 bytes)
        data.extend_from_slice(&self.properties.to_le_bytes());
        data.extend_from_slice(&self.rows.to_le_bytes());
        data.extend_from_slice(&self.cols.to_le_bytes());
        data.extend_from_slice(&self.cell_spacing.to_le_bytes());
        data.extend_from_slice(&self.left_margin.to_le_bytes());
        data.extend_from_slice(&self.right_margin.to_le_bytes());
        data.extend_from_slice(&self.top_margin.to_le_bytes());
        data.extend_from_slice(&self.bottom_margin.to_le_bytes());

        // Cells data (in row order)
        let sorted_cells = self.cells_by_row();
        for cell in sorted_cells {
            data.extend_from_slice(&cell.list_header_id.to_le_bytes());
            data.extend_from_slice(&cell.col_span.to_le_bytes());
            data.extend_from_slice(&cell.row_span.to_le_bytes());
            data.extend_from_slice(&cell.width.to_le_bytes());
            data.extend_from_slice(&cell.height.to_le_bytes());
            data.extend_from_slice(&cell.left_margin.to_le_bytes());
            data.extend_from_slice(&cell.right_margin.to_le_bytes());
            data.extend_from_slice(&cell.top_margin.to_le_bytes());
            data.extend_from_slice(&cell.bottom_margin.to_le_bytes());
            data.extend_from_slice(&cell.border_fill_id.to_le_bytes());
            data.extend_from_slice(&cell.text_width.to_le_bytes());

            // Field name (length + UTF-16LE string)
            let name_utf16: Vec<u16> = cell.field_name.encode_utf16().collect();
            data.extend_from_slice(&(name_utf16.len() as u16).to_le_bytes());
            for ch in name_utf16 {
                data.extend_from_slice(&ch.to_le_bytes());
            }
        }

        data
    }
}

impl TableCell {
    /// Parse cell data from a LIST_HEADER record in table context.
    /// HWP 5.0 cell LIST_HEADER: nPara(u16) + properties(u32) + colAddr(u16) + rowAddr(u16)
    /// + colSpan(u16) + rowSpan(u16) + width(u32) + height(u32) + margins(u16×4) + borderFillId(u16)
    pub fn from_list_header_record(
        record: &crate::parser::record::Record,
    ) -> crate::error::Result<Self> {
        let mut reader = record.data_reader();

        if reader.remaining() < 32 {
            return Err(crate::error::HwpError::ParseError(format!(
                "Cell LIST_HEADER too small: {} bytes",
                reader.remaining()
            )));
        }

        let _n_para = reader.read_u16()?;
        let _properties = reader.read_u32()?;
        let col_addr = reader.read_u16()?;
        let row_addr = reader.read_u16()?;
        let col_span = reader.read_u16()?;
        let row_span = reader.read_u16()?;
        let width = reader.read_u32()?;
        let height = reader.read_u32()?;
        let left_margin = reader.read_u16()?;
        let right_margin = reader.read_u16()?;
        let top_margin = reader.read_u16()?;
        let bottom_margin = reader.read_u16()?;
        let border_fill_id = reader.read_u16()?;

        Ok(Self {
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
            text_width: width.saturating_sub((left_margin + right_margin) as u32),
            field_name: String::new(),
            paragraph_list_id: None,
            cell_address: (row_addr, col_addr),
        })
    }

    pub fn new_default(width: u32, height: u32) -> Self {
        Self {
            list_header_id: 0,
            col_span: 1,
            row_span: 1,
            width,
            height,
            left_margin: 100,
            right_margin: 100,
            top_margin: 100,
            bottom_margin: 100,
            border_fill_id: 0,
            text_width: width.saturating_sub(200),
            field_name: format!("Cell{}x{}", width / 100, height / 100),
            paragraph_list_id: None,
            cell_address: (0, 0),
        }
    }
}

impl Table {
    /// Parse TABLE record from HWP 5.0 body text (tag 0x4D)
    /// Note: cell data comes from LIST_HEADER records, not from this record.
    pub fn from_record(record: &crate::parser::record::Record) -> crate::error::Result<Self> {
        let mut reader = record.data_reader();

        if reader.remaining() < 10 {
            return Err(crate::error::HwpError::ParseError(format!(
                "Table record too small: {} bytes",
                reader.remaining()
            )));
        }

        let properties = reader.read_u32()?;
        let rows = reader.read_u16()?;
        let cols = reader.read_u16()?;
        let cell_spacing = reader.read_u16()?;

        // HWP 5.0 margins are u16, not i32
        let left_margin = if reader.remaining() >= 2 {
            reader.read_u16()? as i32
        } else {
            0
        };
        let right_margin = if reader.remaining() >= 2 {
            reader.read_u16()? as i32
        } else {
            0
        };
        let top_margin = if reader.remaining() >= 2 {
            reader.read_u16()? as i32
        } else {
            0
        };
        let bottom_margin = if reader.remaining() >= 2 {
            reader.read_u16()? as i32
        } else {
            0
        };

        // Remaining data: row sizes (rows * u16), border fill ID, etc.
        // Cell data comes from LIST_HEADER records, not parsed here.

        Ok(Self {
            properties,
            rows,
            cols,
            cell_spacing,
            left_margin,
            right_margin,
            top_margin,
            bottom_margin,
            cells: Vec::new(), // Populated later from LIST_HEADER records
        })
    }
}
