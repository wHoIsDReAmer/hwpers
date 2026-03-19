/// Debug tool: create a table HWP, re-read it, and hex-dump LIST_HEADER records
/// to determine the correct byte layout for cell parsing.
use hwpers::parser::record::{HwpTag, Record};
use hwpers::reader::{CfbReader, StreamReader};
use hwpers::utils::compression::decompress_stream;
use hwpers::HwpWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Step 1: Create an HWP file with a known table
    println!("=== Creating test HWP with 3x2 table ===");
    let mut writer = HwpWriter::new();
    writer.add_simple_table(&[
        vec!["A1", "B1"],
        vec!["A2", "B2"],
        vec!["A3", "B3"],
    ])?;
    writer.save_to_file("/tmp/table_test.hwp")?;
    println!("Saved /tmp/table_test.hwp\n");

    // Step 2: Read it back and dump records
    println!("=== Re-reading and analyzing records ===\n");
    let mut reader = CfbReader::from_file("/tmp/table_test.hwp")?;
    let header_data = reader.read_stream("FileHeader")?;
    let header = hwpers::parser::header::FileHeader::parse(header_data)?;

    let section_data = reader.read_stream("BodyText/Section0")?;
    let data = if header.is_compressed() {
        decompress_stream(&section_data)?
    } else {
        section_data
    };

    let mut stream = StreamReader::new(data);
    let mut record_idx = 0;
    let mut in_table = false;
    let mut last_tag: Option<HwpTag> = None;

    while stream.remaining() >= 4 {
        let pos = stream.position();
        let record = match Record::parse(&mut stream) {
            Ok(r) => r,
            Err(e) => {
                println!("Parse error at 0x{:04X}: {}", pos, e);
                break;
            }
        };

        let tag = HwpTag::from_u16(record.tag_id());
        let tag_name = tag
            .map(|t| format!("{:?}", t))
            .unwrap_or_else(|| format!("0x{:04X}", record.tag_id()));

        // Track table context
        match tag {
            Some(HwpTag::PageHide) | Some(HwpTag::Table) => {
                // 0x4D = TABLE in body text
                in_table = true;
                println!(
                    "#{:3} [pos=0x{:04X}] TAG={} (TABLE) level={} size={}",
                    record_idx, pos, tag_name, record.header.level, record.data.len()
                );
                hex_dump(&record.data);
                println!();
            }
            Some(HwpTag::HiddenComment) | Some(HwpTag::ListHeader) => {
                // 0x48 / 0x56 = LIST_HEADER
                println!(
                    "#{:3} [pos=0x{:04X}] TAG={} (LIST_HEADER) level={} size={} in_table={}",
                    record_idx, pos, tag_name, record.header.level, record.data.len(), in_table
                );
                hex_dump(&record.data);

                // Try parsing as cell with current format (u16 nPara)
                if in_table {
                    try_parse_cell_u16(&record.data);
                    try_parse_cell_i32(&record.data);
                }
                println!();
            }
            Some(HwpTag::SectionDefine) | Some(HwpTag::ParaHeader) => {
                // 0x42 / 0x50 = PARA_HEADER
                println!(
                    "#{:3} [pos=0x{:04X}] TAG={} (PARA_HEADER) level={} size={}",
                    record_idx, pos, tag_name, record.header.level, record.data.len()
                );
            }
            Some(HwpTag::ColumnDefine) | Some(HwpTag::ParaText) => {
                // 0x43 / 0x51 = PARA_TEXT
                let text_preview = String::from_utf16_lossy(
                    &record
                        .data
                        .chunks(2)
                        .take(20)
                        .map(|c| {
                            if c.len() == 2 {
                                u16::from_le_bytes([c[0], c[1]])
                            } else {
                                0
                            }
                        })
                        .collect::<Vec<_>>(),
                );
                println!(
                    "#{:3} [pos=0x{:04X}] TAG={} (PARA_TEXT) text=\"{}\"",
                    record_idx, pos, tag_name, text_preview.trim_end_matches('\0')
                );
            }
            _ => {
                println!(
                    "#{:3} [pos=0x{:04X}] TAG={} level={} size={}",
                    record_idx, pos, tag_name, record.header.level, record.data.len()
                );
            }
        }

        last_tag = tag;
        record_idx += 1;
    }

    Ok(())
}

fn hex_dump(data: &[u8]) {
    for (i, chunk) in data.chunks(16).enumerate().take(4) {
        print!("    {:04X}: ", i * 16);
        for &b in chunk {
            print!("{:02X} ", b);
        }
        for _ in chunk.len()..16 {
            print!("   ");
        }
        print!(" | ");
        for &b in chunk {
            if b.is_ascii_graphic() || b == b' ' {
                print!("{}", b as char);
            } else {
                print!(".");
            }
        }
        println!();
    }
    if data.len() > 64 {
        println!("    ... ({} more bytes)", data.len() - 64);
    }
}

/// Try parsing cell with nPara as u16 (current code)
fn try_parse_cell_u16(data: &[u8]) {
    if data.len() < 32 {
        println!("    [u16 parse] too small ({} bytes)", data.len());
        return;
    }
    let n_para = u16::from_le_bytes([data[0], data[1]]);
    let props = u32::from_le_bytes([data[2], data[3], data[4], data[5]]);
    let col = u16::from_le_bytes([data[6], data[7]]);
    let row = u16::from_le_bytes([data[8], data[9]]);
    let colspan = u16::from_le_bytes([data[10], data[11]]);
    let rowspan = u16::from_le_bytes([data[12], data[13]]);
    let w = u32::from_le_bytes([data[14], data[15], data[16], data[17]]);
    let h = u32::from_le_bytes([data[18], data[19], data[20], data[21]]);
    println!(
        "    [u16 parse] nPara={} props=0x{:08X} addr=({},{}) span={}x{} w={} h={}",
        n_para, props, row, col, colspan, rowspan, w, h
    );
}

/// Try parsing cell with nPara as i32 (standard ListHeader)
fn try_parse_cell_i32(data: &[u8]) {
    if data.len() < 34 {
        println!("    [i32 parse] too small ({} bytes)", data.len());
        return;
    }
    let n_para = i32::from_le_bytes([data[0], data[1], data[2], data[3]]);
    let props = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
    let col = u16::from_le_bytes([data[8], data[9]]);
    let row = u16::from_le_bytes([data[10], data[11]]);
    let colspan = u16::from_le_bytes([data[12], data[13]]);
    let rowspan = u16::from_le_bytes([data[14], data[15]]);
    let w = u32::from_le_bytes([data[16], data[17], data[18], data[19]]);
    let h = u32::from_le_bytes([data[20], data[21], data[22], data[23]]);
    println!(
        "    [i32 parse] nPara={} props=0x{:08X} addr=({},{}) span={}x{} w={} h={}",
        n_para, props, row, col, colspan, rowspan, w, h
    );
}
