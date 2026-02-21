use calamine::{open_workbook, Reader, Xlsx, Data};
use rust_xlsxwriter::Workbook;
use serde::{Serialize, Deserialize};
use std::collections::HashMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnInfo {
    pub index: usize,
    pub code: String,
    pub title: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ColumnMapping {
    pub source_indices: Vec<usize>,
    pub operation: String, // "concat", "add", "subtract", "multiply", "divide"
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ConvertedRow {
    pub recipient_name: String,
    pub recipient_phone: String,
    pub delivery_address: String,
    pub product_name: String,
    pub product_spec: String,
    pub quantity: String,
    pub remarks: String,
    pub group_id: usize,
}

#[derive(Debug, Serialize)]
pub struct ConversionResult {
    pub rows: Vec<ConvertedRow>,
    pub total_rows: usize,
    pub has_duplicates: bool,
    pub duplicate_count: usize,
}

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::String(s) => s.clone(),
        Data::Float(f) => {
            // 手机号等数字去掉小数点
            if *f == f.trunc() {
                format!("{}", *f as i64)
            } else {
                format!("{}", f)
            }
        }
        Data::Int(i) => format!("{}", i),
        Data::Bool(b) => format!("{}", b),
        Data::DateTime(dt) => format!("{}", dt),
        Data::DateTimeIso(s) => s.clone(),
        Data::DurationIso(s) => s.clone(),
        Data::Error(e) => format!("{:?}", e),
        Data::Empty => String::new(),
    }
}

fn get_cell(row: &[Data], index: usize) -> String {
    row.get(index).map(cell_to_string).unwrap_or_default()
}

fn index_to_column_code(index: usize) -> String {
    let mut code = String::new();
    let mut n = index;
    loop {
        code.insert(0, (b'A' + (n % 26) as u8) as char);
        if n < 26 { break; }
        n = n / 26 - 1;
    }
    code
}

pub fn read_columns(path: &str) -> Result<Vec<ColumnInfo>, String> {
    println!("[DEBUG] read_columns called with path: {}", path);
    let mut workbook: Xlsx<_> = open_workbook(path).map_err(|e| format!("无法打开文件: {}", e))?;
    let sheet_names = workbook.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("文件中没有工作表".to_string());
    }
    let range = workbook.worksheet_range(&sheet_names[0]).map_err(|e| format!("无法读取工作表: {}", e))?;
    
    let mut columns = Vec::new();
    if let Some(first_row) = range.rows().next() {
        for (i, cell) in first_row.iter().enumerate() {
            columns.push(ColumnInfo {
                index: i,
                code: index_to_column_code(i),
                title: cell_to_string(cell),
            });
        }
    }
    println!("[DEBUG] read_columns returning {} columns", columns.len());
    Ok(columns)
}

fn apply_operation(values: Vec<String>, operation: &str) -> String {
    if values.is_empty() { return String::new(); }
    
    match operation {
        "concat" => values.join(""),
        "add" | "subtract" | "multiply" | "divide" => {
            let nums: Vec<f64> = values.iter().filter_map(|v| v.parse().ok()).collect();
            if nums.is_empty() { return String::new(); }
            
            let result = match operation {
                "add" => nums.iter().sum::<f64>(),
                "subtract" => nums.iter().skip(1).fold(nums[0], |acc, &x| acc - x),
                "multiply" => nums.iter().product::<f64>(),
                "divide" => nums.iter().skip(1).fold(nums[0], |acc, &x| if x != 0.0 { acc / x } else { acc }),
                _ => 0.0,
            };
            if result == result.trunc() { format!("{}", result as i64) } else { format!("{}", result) }
        }
        _ => values.join(""),
    }
}

pub fn read_and_convert_with_mapping(
    path: &str,
    mappings: HashMap<String, ColumnMapping>,
) -> Result<ConversionResult, String> {
    let mut workbook: Xlsx<_> = open_workbook(path).map_err(|e| format!("无法打开文件: {}", e))?;
    let sheet_names = workbook.sheet_names().to_vec();
    if sheet_names.is_empty() {
        return Err("文件中没有工作表".to_string());
    }
    let range = workbook.worksheet_range(&sheet_names[0]).map_err(|e| format!("无法读取工作表: {}", e))?;

    let mut rows = Vec::new();

    for (i, row) in range.rows().enumerate() {
        if i == 0 { continue; }

        let mut converted = ConvertedRow {
            recipient_name: String::new(),
            recipient_phone: String::new(),
            delivery_address: String::new(),
            product_name: String::new(),
            product_spec: String::new(),
            quantity: "1".to_string(),
            remarks: String::new(),
            group_id: 0,
        };

        for (field, mapping) in &mappings {
            let values: Vec<String> = mapping.source_indices.iter().map(|&idx| get_cell(row, idx)).collect();
            let result = apply_operation(values, &mapping.operation);
            
            match field.as_str() {
                "recipient_name" => converted.recipient_name = result,
                "recipient_phone" => converted.recipient_phone = result,
                "delivery_address" => converted.delivery_address = result,
                "product_name" => converted.product_name = result,
                "product_spec" => converted.product_spec = result,
                "quantity" => converted.quantity = result,
                "remarks" => converted.remarks = result,
                _ => {}
            }
        }

        if converted.recipient_name.is_empty() && converted.recipient_phone.is_empty() && converted.delivery_address.is_empty() {
            continue;
        }

        rows.push(converted);
    }

    // 按 (name, phone, address) 排序
    rows.sort_by(|a, b| {
        (&a.recipient_name, &a.recipient_phone, &a.delivery_address)
            .cmp(&(&b.recipient_name, &b.recipient_phone, &b.delivery_address))
    });

    let mut key_count: HashMap<(String, String, String), usize> = HashMap::new();
    for row in &rows {
        let key = (row.recipient_name.clone(), row.recipient_phone.clone(), row.delivery_address.clone());
        *key_count.entry(key).or_insert(0) += 1;
    }

    let mut group_map: HashMap<(String, String, String), usize> = HashMap::new();
    let mut next_group_id = 1usize;
    let mut duplicate_count = 0usize;

    for row in &mut rows {
        let key = (row.recipient_name.clone(), row.recipient_phone.clone(), row.delivery_address.clone());
        if let Some(&count) = key_count.get(&key) {
            if count >= 2 {
                let gid = *group_map.entry(key).or_insert_with(|| {
                    let id = next_group_id;
                    next_group_id += 1;
                    duplicate_count += count;
                    id
                });
                row.group_id = gid;
            }
        }
    }

    let has_duplicates = duplicate_count > 0;
    let total_rows = rows.len();
    Ok(ConversionResult { rows, total_rows, has_duplicates, duplicate_count })
}

pub fn read_and_convert(path: &str) -> Result<ConversionResult, String> {
    let mut mappings = HashMap::new();
    mappings.insert("recipient_name".to_string(), ColumnMapping { source_indices: vec![64], operation: "concat".to_string() });
    mappings.insert("recipient_phone".to_string(), ColumnMapping { source_indices: vec![65], operation: "concat".to_string() });
    mappings.insert("delivery_address".to_string(), ColumnMapping { source_indices: vec![69], operation: "concat".to_string() });
    mappings.insert("product_name".to_string(), ColumnMapping { source_indices: vec![82], operation: "concat".to_string() });
    mappings.insert("remarks".to_string(), ColumnMapping { source_indices: vec![0], operation: "concat".to_string() });
    
    read_and_convert_with_mapping(path, mappings)
}

pub fn merge_duplicates(rows: &[ConvertedRow]) -> Vec<ConvertedRow> {
    // 按 (name, phone, address) 分组，保持原有顺序
    let mut group_order: Vec<(String, String, String)> = Vec::new();
    let mut groups: HashMap<(String, String, String), Vec<&ConvertedRow>> = HashMap::new();

    for row in rows {
        let key = (row.recipient_name.clone(), row.recipient_phone.clone(), row.delivery_address.clone());
        if !groups.contains_key(&key) {
            group_order.push(key.clone());
        }
        groups.entry(key).or_default().push(row);
    }

    let mut result = Vec::new();

    for key in &group_order {
        let group = &groups[key];
        if group.len() == 1 {
            let mut row = group[0].clone();
            row.group_id = 0;
            result.push(row);
        } else {
            // 合并商品名称（去重去空，用；拼接）
            let mut product_names: Vec<String> = Vec::new();
            let mut product_specs: Vec<String> = Vec::new();
            let mut total_quantity: usize = 0;
            let mut remarks_list: Vec<String> = Vec::new();

            for r in group {
                if !r.product_name.is_empty() && !product_names.contains(&r.product_name) {
                    product_names.push(r.product_name.clone());
                }
                if !r.product_spec.is_empty() && !product_specs.contains(&r.product_spec) {
                    product_specs.push(r.product_spec.clone());
                }
                let qty: usize = r.quantity.parse().unwrap_or(1);
                total_quantity += qty;
                if !r.remarks.is_empty() && !remarks_list.contains(&r.remarks) {
                    remarks_list.push(r.remarks.clone());
                }
            }

            // 数量：如果有多个商品名，每个商品各为1；否则求和
            let quantity_str = if product_names.len() > 1 {
                product_names.iter().map(|_| "1").collect::<Vec<_>>().join("；")
            } else {
                total_quantity.to_string()
            };

            result.push(ConvertedRow {
                recipient_name: key.0.clone(),
                recipient_phone: key.1.clone(),
                delivery_address: key.2.clone(),
                product_name: product_names.join("；"),
                product_spec: product_specs.join("；"),
                quantity: quantity_str,
                remarks: remarks_list.join("；"),
                group_id: 0,
            });
        }
    }

    result
}

pub fn write_output(rows: &[ConvertedRow], output_path: &str) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let sheet = workbook.add_worksheet();

    sheet.set_name("工作表1").map_err(|e| format!("设置工作表名失败: {}", e))?;

    // 写入表头
    let headers = [
        "收件人姓名（必填）",
        "收件人手机号（必填）",
        "收货地址（必填）",
        "商品名称(必填) -- 多商品用\u{201c}；\u{201d}隔开",
        "商品规格(非必填) -- 多商品用\u{201c}；\u{201d}隔开",
        "商品数量(必填) -- 多商品用\u{201c}；\u{201d}隔开",
        "备注（非必填）",
    ];

    for (col, header) in headers.iter().enumerate() {
        sheet
            .write_string(0, col as u16, *header)
            .map_err(|e| format!("写入表头失败: {}", e))?;
    }

    // 写入数据行
    for (row_idx, row) in rows.iter().enumerate() {
        let r = (row_idx + 1) as u32; // 数据从第2行开始
        sheet.write_string(r, 0, &row.recipient_name).map_err(|e| format!("写入失败: {}", e))?;
        sheet.write_string(r, 1, &row.recipient_phone).map_err(|e| format!("写入失败: {}", e))?;
        sheet.write_string(r, 2, &row.delivery_address).map_err(|e| format!("写入失败: {}", e))?;
        sheet.write_string(r, 3, &row.product_name).map_err(|e| format!("写入失败: {}", e))?;
        sheet.write_string(r, 4, &row.product_spec).map_err(|e| format!("写入失败: {}", e))?;
        sheet.write_string(r, 5, &row.quantity).map_err(|e| format!("写入失败: {}", e))?;
        sheet.write_string(r, 6, &row.remarks).map_err(|e| format!("写入失败: {}", e))?;
    }

    workbook
        .save(output_path)
        .map_err(|e| format!("保存文件失败: {}", e))?;

    Ok(())
}
