use crate::model::{
    ApartmentRecord, DepartmentRecord, GradeRecord, ProcessedRecord, ReportDataRecord,
};
use anyhow::Result;
use csv::ReaderBuilder;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Image, Workbook, Worksheet};
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    path::{Path, PathBuf},
    sync::LazyLock,
};

static GRADE_MAP: LazyLock<HashMap<(u8, u8), (String, String)>> =
    LazyLock::new(|| load_grade_data("assets/grade.csv").unwrap());

static APT_MAP: LazyLock<HashMap<(u8, u8), String>> =
    LazyLock::new(|| load_apt_data("assets/apt.csv").unwrap());

static DPT_MAP: LazyLock<HashMap<(u8, String), (String, u8)>> =
    LazyLock::new(|| load_dept_data("assets/dpt.csv").unwrap());

static ALL_MANAGERS: LazyLock<Vec<(u8, u8, String)>> =
    LazyLock::new(|| get_all_managers("assets/apt.csv").unwrap());

fn output_path(input: &Path, output: Option<PathBuf>) -> PathBuf {
    output.unwrap_or_else(|| {
        let mut out: PathBuf = input.into();
        out.set_extension("xlsx");
        out
    })
}

struct ReportFormats {
    title: Format,
    header: Format,
    cell: Format,
    left_align: Format,
    center_bold: Format,
    left_text: Format,
}

impl ReportFormats {
    fn new() -> Self {
        Self {
            title: Format::new()
                .set_bold()
                .set_font_size(18)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter),
            header: Format::new()
                .set_bold()
                .set_border(FormatBorder::Thin)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_text_wrap(),
            cell: Format::new()
                .set_border(FormatBorder::Thin)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_text_wrap(),
            left_align: Format::new()
                .set_align(FormatAlign::Left)
                .set_border(FormatBorder::Thin)
                .set_bold()
                .set_align(FormatAlign::VerticalCenter),
            center_bold: Format::new()
                .set_align(FormatAlign::Center)
                .set_border(FormatBorder::Thin)
                .set_bold()
                .set_align(FormatAlign::VerticalCenter),
            left_text: Format::new()
                .set_align(FormatAlign::Left)
                .set_border(FormatBorder::Thin)
                .set_text_wrap()
                .set_align(FormatAlign::VerticalCenter),
        }
    }
}

const RULES: &str = "宿舍卫生:宿舍卫生验评满分10分\n1.宿舍床铺被子叠放整齐(此项不合格每人扣1分)\n2.床单平整(此项不合格每人扣1分)\n3.无多余杂物(如衣物、书本、零食)此项不合格每人扣1分)\n4.簸箕内清理干净(此项不合格每人扣1分)";

fn grade_name(grade: u8) -> &'static str {
    match grade {
        1 => "高一",
        2 => "高二",
        3 => "高三",
        _ => "",
    }
}

fn apt_display_name(apt: u8) -> String {
    format!("{}号公寓", if apt == 1 { "一" } else { "二" })
}

fn compute_ranks<K: Clone + Eq + std::hash::Hash>(totals: &[(K, i32)]) -> HashMap<K, i32> {
    let mut rank_map = HashMap::new();
    if totals.is_empty() {
        return rank_map;
    }
    let mut cur_rank = 1;
    let mut prev_score = totals[0].1;
    rank_map.insert(totals[0].0.clone(), cur_rank);
    for (key, score) in totals.iter().skip(1) {
        if *score != prev_score {
            cur_rank += 1;
            prev_score = *score;
        }
        rank_map.insert(key.clone(), cur_rank);
    }
    rank_map
}

fn write_report_header(
    ws: &mut Worksheet,
    start_row: u32,
    reporter: &str,
    date: &str,
    time: &str,
    fmt: &ReportFormats,
) -> Result<u32> {
    // 设置标题行高度（像素），logo 高度与之匹配
    const TITLE_ROW_HEIGHT: f64 = 30.0;
    const LOGO_HEIGHT: u32 = 40; // 像素，约等于行高

    ws.set_row_height(start_row, TITLE_ROW_HEIGHT)?;
    ws.merge_range(
        start_row,
        0,
        start_row,
        8,
        "高中部宿舍卫生验评通报总结",
        &fmt.title,
    )?;
    let image = Image::new("assets/logo.png")?
        .set_height(LOGO_HEIGHT)
        .set_width(LOGO_HEIGHT); // 保持正方形
    // 设置 logo 在单元格内垂直居中的偏移量
    ws.insert_image_with_offset(start_row, 0, &image, 0, 5)?;
    let r = start_row + 1;
    ws.merge_range(
        r,
        0,
        r,
        4,
        &format!("汇报人: {}", reporter),
        &fmt.left_align,
    )?;
    ws.merge_range(r, 5, r, 7, "验评对象: 高一、高二、高三", &fmt.center_bold)?;
    ws.write_string_with_format(r, 8, format!("日期: {}", date), &fmt.center_bold)?;
    let r = r + 1;
    ws.write_string_with_format(r, 0, "验评部门", &fmt.center_bold)?;
    ws.merge_range(r, 1, r, 8, "校办公室", &fmt.cell)?;
    let r = r + 1;
    ws.write_string_with_format(r, 0, "验评项目", &fmt.center_bold)?;
    ws.merge_range(r, 1, r, 8, "高一高二高三男生宿舍卫生", &fmt.cell)?;
    let r = r + 1;
    ws.write_string_with_format(r, 0, "验评时间", &fmt.center_bold)?;
    ws.merge_range(r, 1, r, 8, time, &fmt.cell)?;
    let r = r + 1;
    ws.write_string_with_format(r, 0, "验评细则", &fmt.center_bold)?;
    ws.merge_range(r, 1, r, 8, RULES, &fmt.left_text)?;
    ws.set_row_height(r, 80)?;
    Ok(r + 1)
}

fn merge_or_write_str(
    ws: &mut Worksheet,
    start: u32,
    end: u32,
    col: u16,
    val: &str,
    fmt: &Format,
) -> Result<()> {
    if end > start {
        ws.merge_range(start, col, end, col, val, fmt)?;
    } else {
        ws.write_string_with_format(start, col, val, fmt)?;
    }
    Ok(())
}

fn merge_or_write_num(
    ws: &mut Worksheet,
    start: u32,
    end: u32,
    col: u16,
    val: f64,
    fmt: &Format,
) -> Result<()> {
    if end > start {
        ws.merge_range(start, col, end, col, &val.to_string(), fmt)?;
    } else {
        ws.write_number_with_format(start, col, val, fmt)?;
    }
    Ok(())
}

fn write_table1_headers(ws: &mut Worksheet, row: u32, fmt: &Format) -> Result<()> {
    let headers = [
        "公寓",
        "级部",
        "班主任",
        "宿舍管理员",
        "宿舍号",
        "扣分原因",
        "扣分",
        "总扣分",
        "排名",
    ];
    for (i, h) in headers.iter().enumerate() {
        ws.write_string_with_format(row, i as u16, *h, fmt)?;
    }
    Ok(())
}

fn write_table2_headers(ws: &mut Worksheet, row: u32, fmt: &Format) -> Result<()> {
    ws.write_string_with_format(row, 0, "公寓", fmt)?;
    ws.write_string_with_format(row, 1, "宿舍管理员", fmt)?;
    ws.write_string_with_format(row, 2, "宿舍号", fmt)?;
    ws.merge_range(row, 3, row, 4, "扣分原因", fmt)?;
    ws.write_string_with_format(row, 5, "扣分", fmt)?;
    ws.merge_range(row, 6, row, 7, "总扣分", fmt)?;
    ws.write_string_with_format(row, 8, "排名", fmt)?;
    Ok(())
}

fn set_column_widths(ws: &mut Worksheet) -> Result<()> {
    let widths = [12, 12, 12, 10, 10, 18, 8, 8, 8];
    for (col, w) in widths.iter().enumerate() {
        ws.set_column_width(col as u16, *w)?;
    }
    Ok(())
}

struct Apt2AState {
    in_both: bool,
    in_apt1_only: bool,
    in_apt2_only: bool,
    in_neither: bool,
    start_row: Option<u32>,
    end_row: Option<u32>,
}

impl Apt2AState {
    fn new(data: &[ProcessedRecord]) -> Self {
        let mut has_records: HashMap<u8, bool> = HashMap::new();
        for r in data {
            if r.grade == 2 && r.dept == "A" {
                has_records.insert(r.apartment, true);
            }
        }
        Self {
            in_both: has_records.contains_key(&1) && has_records.contains_key(&2),
            in_apt1_only: has_records.contains_key(&1) && !has_records.contains_key(&2),
            in_apt2_only: has_records.contains_key(&2) && !has_records.contains_key(&1),
            in_neither: has_records.is_empty(),
            start_row: None,
            end_row: None,
        }
    }

    fn should_show_in_apt(&self, apt: u8) -> bool {
        self.in_both
            || (self.in_apt1_only && apt == 1)
            || (self.in_apt2_only && apt == 2)
            || (self.in_neither && apt == 1)
    }
}

fn write_dorm_row_table1(
    ws: &mut Worksheet,
    row: u32,
    r: &ProcessedRecord,
    fmt: &Format,
) -> Result<()> {
    ws.write_string_with_format(row, 2, &r.teacher, fmt)?;
    ws.write_string_with_format(row, 3, &r.manager, fmt)?;
    ws.write_string_with_format(row, 4, format!("{}宿舍", r.dorm), fmt)?;
    ws.write_string_with_format(row, 5, &r.reason, fmt)?;
    ws.write_number_with_format(row, 6, r.deduction as f64, fmt)?;
    Ok(())
}

fn write_empty_dept_row(
    ws: &mut Worksheet,
    row: u32,
    dept_display: &str,
    rank: i32,
    fmt: &Format,
) -> Result<()> {
    ws.write_string_with_format(row, 1, dept_display, fmt)?;
    for col in 2..=7 {
        ws.write_string_with_format(row, col, "/", fmt)?;
    }
    ws.write_number_with_format(row, 8, rank as f64, fmt)?;
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_dept_group(
    ws: &mut Worksheet,
    row: &mut u32,
    grade: u8,
    dept: &str,
    records: &[&ProcessedRecord],
    global_rank_map: &HashMap<(u8, String), i32>,
    dpt_map: &HashMap<(u8, String), (String, u8)>,
    apt2a: &mut Apt2AState,
    fmt: &Format,
) -> Result<()> {
    let leader = dpt_map
        .get(&(grade, dept.to_string()))
        .map(|(l, _)| l.clone())
        .unwrap_or_default();
    let dept_display = format!("{}{}部\n({})", grade_name(grade), dept, leader);
    let grp_start = *row;
    let is_2a = grade == 2 && dept == "A";

    if is_2a && apt2a.in_both && apt2a.start_row.is_none() {
        apt2a.start_row = Some(*row);
    }

    let rank = *global_rank_map
        .get(&(grade, dept.to_string()))
        .unwrap_or(&0);

    if records.is_empty() {
        write_empty_dept_row(ws, *row, &dept_display, rank, fmt)?;
        *row += 1;
    } else {
        let mut sorted: Vec<_> = records.to_vec();
        sorted.sort_by_key(|r| r.dorm);
        let total: i32 = sorted.iter().map(|r| r.deduction).sum();

        for (idx, r) in sorted.iter().enumerate() {
            write_dorm_row_table1(ws, grp_start + idx as u32, r, fmt)?;
        }
        *row += sorted.len() as u32;

        if is_2a && apt2a.in_both {
            apt2a.end_row = Some(*row - 1);
        }

        if !(is_2a && apt2a.in_both) {
            let end = *row - 1;
            merge_or_write_str(ws, grp_start, end, 1, &dept_display, fmt)?;
            merge_or_write_str(ws, grp_start, end, 7, &total.to_string(), fmt)?;
            merge_or_write_num(ws, grp_start, end, 8, rank as f64, fmt)?;
        }
    }
    Ok(())
}

fn write_class_group(
    ws: &mut Worksheet,
    row: &mut u32,
    class_num: u8,
    records: &[&ProcessedRecord],
    class_rank_map: &HashMap<u8, i32>,
    fmt: &Format,
) -> Result<()> {
    if records.is_empty() {
        return Ok(());
    }

    let mut sorted: Vec<_> = records.to_vec();
    sorted.sort_by_key(|r| r.dorm);
    let total: i32 = sorted.iter().map(|r| r.deduction).sum();
    let rank = *class_rank_map.get(&class_num).unwrap_or(&0);
    let class_display = format!("{}班", class_num);
    let grp_start = *row;

    for (idx, r) in sorted.iter().enumerate() {
        write_dorm_row_table1(ws, grp_start + idx as u32, r, fmt)?;
    }
    *row += sorted.len() as u32;

    let end = *row - 1;
    merge_or_write_str(ws, grp_start, end, 1, &class_display, fmt)?;
    merge_or_write_str(ws, grp_start, end, 7, &total.to_string(), fmt)?;
    merge_or_write_num(ws, grp_start, end, 8, rank as f64, fmt)?;
    Ok(())
}

fn write_table1(
    ws: &mut Worksheet,
    start_row: u32,
    data: &[ProcessedRecord],
    dpt_map: &HashMap<(u8, String), (String, u8)>,
    fmt: &ReportFormats,
) -> Result<u32> {
    write_table1_headers(ws, start_row, &fmt.header)?;
    let mut row = start_row + 1;

    // 公寓列表改为从级部配置中推导，而不是仅从实际数据中推导，
    // 这样即使当天没有任何记录，也会为所有配置过的公寓生成表格结构。
    let mut apartments: Vec<u8> = dpt_map
        .values()
        .map(|(_, apt)| *apt)
        .collect::<HashSet<_>>()
        .into_iter()
        .collect();
    apartments.sort_by(|a, b| b.cmp(a));

    // Global rankings
    let mut all_dept_groups: HashMap<(u8, String), Vec<&ProcessedRecord>> = HashMap::new();
    for (grade, dept) in dpt_map.keys() {
        all_dept_groups.entry((*grade, dept.clone())).or_default();
    }
    for r in data {
        if !r.dept.is_empty() {
            all_dept_groups
                .entry((r.grade, r.dept.clone()))
                .or_default()
                .push(r);
        }
    }
    let mut all_dept_totals: Vec<((u8, String), i32)> = all_dept_groups
        .iter()
        .map(|(k, v)| (k.clone(), v.iter().map(|r| r.deduction).sum()))
        .collect();
    all_dept_totals.sort_by(|a, b| b.1.cmp(&a.1));
    let global_rank_map = compute_ranks(&all_dept_totals);

    let mut apt2a = Apt2AState::new(data);

    for apt in &apartments {
        let apt_start = row;
        let mut dept_groups: HashMap<(u8, String), Vec<&ProcessedRecord>> = HashMap::new();
        let mut class_groups: HashMap<u8, Vec<&ProcessedRecord>> = HashMap::new();

        // Initialize departments for this apartment
        for ((grade, dept), (_, default_apt)) in dpt_map.iter() {
            if *grade == 2 && dept == "A" {
                if apt2a.should_show_in_apt(*apt) {
                    dept_groups.entry((*grade, dept.clone())).or_default();
                }
            } else if *default_apt == *apt {
                dept_groups.entry((*grade, dept.clone())).or_default();
            }
        }

        for r in data.iter().filter(|r| r.apartment == *apt) {
            if r.dept.is_empty() {
                class_groups.entry(r.class).or_default().push(r);
            } else {
                dept_groups
                    .entry((r.grade, r.dept.clone()))
                    .or_default()
                    .push(r);
            }
        }

        let mut class_totals: Vec<(u8, i32)> = class_groups
            .iter()
            .map(|(k, v)| (*k, v.iter().map(|r| r.deduction).sum()))
            .collect();
        class_totals.sort_by(|a, b| b.1.cmp(&a.1));
        let class_rank_map = compute_ranks(&class_totals);

        let mut sorted_dept_keys: Vec<_> = dept_groups.keys().cloned().collect();
        sorted_dept_keys.sort_by(|a, b| a.0.cmp(&b.0).then(a.1.cmp(&b.1)));

        let mut sorted_class_keys: Vec<_> = class_groups.keys().cloned().collect();
        sorted_class_keys.sort();

        for (grade, dept) in sorted_dept_keys {
            let records: Vec<_> = dept_groups.get(&(grade, dept.clone())).unwrap().to_vec();
            write_dept_group(
                ws,
                &mut row,
                grade,
                &dept,
                &records,
                &global_rank_map,
                dpt_map,
                &mut apt2a,
                &fmt.cell,
            )?;
        }

        for class_num in sorted_class_keys {
            let records: Vec<_> = class_groups.get(&class_num).unwrap().to_vec();
            write_class_group(
                ws,
                &mut row,
                class_num,
                &records,
                &class_rank_map,
                &fmt.cell,
            )?;
        }

        if row > apt_start {
            merge_or_write_str(
                ws,
                apt_start,
                row - 1,
                0,
                &apt_display_name(*apt),
                &fmt.cell,
            )?;
        }
    }

    // Handle 高二A部 cross-apartment merging
    if apt2a.in_both
        && let (Some(start), Some(end)) = (apt2a.start_row, apt2a.end_row)
    {
        let leader = dpt_map
            .get(&(2, "A".to_string()))
            .map(|(l, _)| l.clone())
            .unwrap_or_default();
        let dept_display = format!("高二A部\n({})", leader);
        let total: i32 = all_dept_groups
            .get(&(2, "A".to_string()))
            .map(|v| v.iter().map(|r| r.deduction).sum())
            .unwrap_or(0);
        let rank = *global_rank_map.get(&(2, "A".to_string())).unwrap_or(&0);
        ws.merge_range(start, 1, end, 1, &dept_display, &fmt.cell)?;
        ws.merge_range(start, 7, end, 7, &total.to_string(), &fmt.cell)?;
        ws.merge_range(start, 8, end, 8, &rank.to_string(), &fmt.cell)?;
    }

    Ok(row)
}

fn write_table2(
    ws: &mut Worksheet,
    start_row: u32,
    data: &[ProcessedRecord],
    all_managers: &[(u8, u8, String)],
    fmt: &ReportFormats,
) -> Result<u32> {
    write_table2_headers(ws, start_row, &fmt.header)?;
    let mut row = start_row + 1;

    let mut mgr_by_apt: HashMap<u8, HashSet<String>> = HashMap::new();
    for (apt, _, name) in all_managers.iter() {
        mgr_by_apt.entry(*apt).or_default().insert(name.clone());
    }
    for r in data {
        mgr_by_apt
            .entry(r.apartment)
            .or_default()
            .insert(r.manager.clone());
    }

    let mut sorted_apts: Vec<u8> = mgr_by_apt.keys().cloned().collect();
    sorted_apts.sort();

    for apt in sorted_apts {
        let mgrs = mgr_by_apt.get(&apt).unwrap();
        let mut mgr_totals: Vec<(String, i32)> = mgrs
            .iter()
            .map(|m| {
                let t: i32 = data
                    .iter()
                    .filter(|r| r.apartment == apt && &r.manager == m)
                    .map(|r| r.deduction)
                    .sum();
                (m.clone(), t)
            })
            .collect();
        mgr_totals.sort_by(|a, b| b.1.cmp(&a.1));
        let rank_map = compute_ranks(&mgr_totals);

        let mut mgr_floors: HashMap<String, u8> = HashMap::new();
        for (a, f, n) in all_managers.iter() {
            if *a == apt {
                let e = mgr_floors.entry(n.clone()).or_insert(*f);
                if *f < *e {
                    *e = *f;
                }
            }
        }

        let mut sorted_mgrs = mgr_totals.clone();
        sorted_mgrs.sort_by_key(|(n, _)| mgr_floors.get(n).cloned().unwrap_or(99));

        let apt_start = row;

        for (mgr, total) in sorted_mgrs {
            let rank = *rank_map.get(&mgr).unwrap();
            let recs: Vec<_> = data
                .iter()
                .filter(|r| r.apartment == apt && r.manager == mgr)
                .collect();
            let mgr_start = row;

            if recs.is_empty() {
                ws.write_string_with_format(row, 1, &mgr, &fmt.cell)?;
                ws.write_string_with_format(row, 2, "/", &fmt.cell)?;
                ws.merge_range(row, 3, row, 4, "/", &fmt.cell)?;
                ws.write_string_with_format(row, 5, "/", &fmt.cell)?;
                ws.merge_range(row, 6, row, 7, "/", &fmt.cell)?;
                ws.write_number_with_format(row, 8, rank as f64, &fmt.cell)?;
                row += 1;
            } else {
                let mut sorted_recs: Vec<_> = recs.iter().collect();
                sorted_recs.sort_by_key(|r| r.dorm);

                for r in &sorted_recs {
                    ws.write_string_with_format(row, 2, format!("{}宿舍", r.dorm), &fmt.cell)?;
                    ws.merge_range(row, 3, row, 4, &r.reason, &fmt.cell)?;
                    ws.write_number_with_format(row, 5, r.deduction as f64, &fmt.cell)?;
                    row += 1;
                }

                if row > mgr_start {
                    let end = row - 1;
                    merge_or_write_str(ws, mgr_start, end, 1, &mgr, &fmt.cell)?;
                    if end > mgr_start {
                        ws.merge_range(mgr_start, 6, end, 7, &total.to_string(), &fmt.cell)?;
                    } else {
                        ws.merge_range(mgr_start, 6, mgr_start, 7, &total.to_string(), &fmt.cell)?;
                    }
                    merge_or_write_num(ws, mgr_start, end, 8, rank as f64, &fmt.cell)?;
                }
            }
        }

        if row > apt_start {
            merge_or_write_str(ws, apt_start, row - 1, 0, &apt_display_name(apt), &fmt.cell)?;
        }
    }

    Ok(row)
}

pub fn generate_report(
    input: PathBuf,
    output: Option<PathBuf>,
    reporter: String,
    date: String,
    time: String,
) -> Result<()> {
    let output_path = output_path(&input, output);
    let processed_data = load_report_data(&input)?;
    let all_managers = &ALL_MANAGERS;
    let dpt_map = &DPT_MAP;

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let fmt = ReportFormats::new();

    // Table 1: Department-based report
    let row = write_report_header(worksheet, 0, &reporter, &date, &time, &fmt)?;
    let row = write_table1(worksheet, row, &processed_data, dpt_map, &fmt)?;

    // Table 2: Manager-based report
    let row = row + 2;
    let row = write_report_header(worksheet, row, &reporter, &date, &time, &fmt)?;
    write_table2(worksheet, row, &processed_data, all_managers, &fmt)?;

    set_column_widths(worksheet)?;
    workbook.save(&output_path)?;
    println!("报告已生成: {}", output_path.display());
    Ok(())
}

fn load_report_data<P: AsRef<Path>>(path: P) -> Result<Vec<ProcessedRecord>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut records = Vec::new();
    for result in rdr.deserialize() {
        let raw_record: ReportDataRecord = result?;
        let dept_info = GRADE_MAP.get(&(raw_record.grade, raw_record.class));
        let floor = (raw_record.dorm / 100) as u8;
        let manager = APT_MAP
            .get(&(raw_record.apartment, floor))
            .cloned()
            .unwrap_or_else(|| "未知".to_string());
        let (dept, teacher) = match dept_info {
            Some((d, t)) => (d.clone(), t.clone()),
            None => ("".to_string(), "未知".to_string()),
        };
        records.push(ProcessedRecord {
            apartment: raw_record.apartment,
            grade: raw_record.grade,
            class: raw_record.class,
            dept,
            teacher,
            manager,
            dorm: raw_record.dorm,
            reason: raw_record.reason,
            deduction: -1,
        });
    }

    Ok(records)
}

fn load_grade_data<P: AsRef<Path>>(path: P) -> Result<HashMap<(u8, u8), (String, String)>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new()
        .has_headers(true)
        .flexible(true)
        .from_reader(file);
    let mut map = HashMap::new();
    for result in rdr.deserialize() {
        let r: GradeRecord = result?;
        map.insert((r.grade, r.class), (r.dept.unwrap_or_default(), r.teacher));
    }
    Ok(map)
}

fn load_apt_data<P: AsRef<Path>>(path: P) -> Result<HashMap<(u8, u8), String>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut map = HashMap::new();
    for result in rdr.deserialize() {
        let r: ApartmentRecord = result?;
        map.insert((r.apartment, r.floor), r.manager);
    }
    Ok(map)
}

fn get_all_managers<P: AsRef<Path>>(path: P) -> Result<Vec<(u8, u8, String)>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut list = Vec::new();
    for result in rdr.deserialize() {
        let r: ApartmentRecord = result?;
        list.push((r.apartment, r.floor, r.manager));
    }
    Ok(list)
}

fn load_dept_data<P: AsRef<Path>>(path: P) -> Result<HashMap<(u8, String), (String, u8)>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut map = HashMap::new();
    for result in rdr.deserialize() {
        let r: DepartmentRecord = result?;
        map.insert((r.grade, r.dept), (r.leader, r.apartment));
    }
    Ok(map)
}
