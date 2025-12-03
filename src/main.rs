use anyhow::Result;
use clap::Parser;
use csv::ReaderBuilder;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Image, Workbook};
use serde::Deserialize;
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::path::Path;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long, default_value = "侯英敏、杨超超、郭静、赵冰、申淑玲")]
    reporter: String,

    #[arg(short, long, default_value = "12月3日")]
    date: String,

    #[arg(short, long, default_value = "下午: 15:20-15:50")]
    time: String,
}

#[derive(Debug, Deserialize)]
struct TestDataRecord {
    #[serde(rename = "年级")]
    grade: u8,
    #[serde(rename = "班级")]
    class: u8,
    #[serde(rename = "公寓")]
    apartment: u8,
    #[serde(rename = "宿舍")]
    dorm: u16,
    #[serde(rename = "原因")]
    reason: String,
}

#[derive(Debug, Deserialize)]
struct NianjiRecord {
    #[serde(rename = "年级")]
    grade: u8,
    #[serde(rename = "级部")]
    dept: Option<String>,
    #[serde(rename = "班级")]
    class: u8,
    #[serde(rename = "班主任")]
    teacher: String,
}

#[derive(Debug, Deserialize)]
struct SusheRecord {
    #[serde(rename = "公寓")]
    apartment: u8,
    #[serde(rename = "楼层")]
    floor: u8,
    #[serde(rename = "宿管")]
    manager: String,
}

struct ProcessedRecord {
    apartment: u8,
    grade: u8,
    dept: String,
    teacher: String,
    manager: String,
    dorm: u16,
    reason: String,
    deduction: i32,
}

fn main() -> Result<()> {
    let args = Args::parse();

    let records = load_test_data("test_data.csv")?;
    let nianji_map = load_nianji_data("assets/nianji.csv")?;
    let sushe_map = load_sushe_data("assets/sushe.csv")?;
    let all_managers = get_all_managers("assets/sushe.csv")?;

    let mut processed_data = Vec::new();
    for r in records {
        let dept_info = nianji_map.get(&(r.grade, r.class));
        let floor = (r.dorm / 100) as u8;
        let manager = sushe_map
            .get(&(r.apartment, floor))
            .cloned()
            .unwrap_or_else(|| "未知".to_string());
        let (dept, teacher) = match dept_info {
            Some((d, t)) => (d.clone(), t.clone()),
            None => ("".to_string(), "未知".to_string()),
        };
        processed_data.push(ProcessedRecord {
            apartment: r.apartment,
            grade: r.grade,
            dept,
            teacher,
            manager,
            dorm: r.dorm,
            reason: r.reason,
            deduction: -1,
        });
    }

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let title_fmt = Format::new()
        .set_bold()
        .set_font_size(18)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);
    let header_fmt = Format::new()
        .set_bold()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap();
    let cell_fmt = Format::new()
        .set_border(FormatBorder::Thin)
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter)
        .set_text_wrap();
    let left_align = Format::new()
        .set_align(FormatAlign::Left)
        .set_border(FormatBorder::Thin)
        .set_bold()
        .set_align(FormatAlign::VerticalCenter);
    let center_bold = Format::new()
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_bold()
        .set_align(FormatAlign::VerticalCenter);
    let left_text = Format::new()
        .set_align(FormatAlign::Left)
        .set_border(FormatBorder::Thin)
        .set_text_wrap()
        .set_align(FormatAlign::VerticalCenter);

    let rules = "宿舍卫生:宿舍卫生验评满分10分\n1.宿舍床铺被子叠放整齐(此项不合格每人扣1分)\n2.床单平整(此项不合格每人扣1分)\n3.无多余杂物(如衣物、书本、零食)此项不合格每人扣1分)\n4.簸箕内清理干净(此项不合格每人扣1分)";

    // Merge range for title first
    worksheet.merge_range(0, 0, 0, 8, "高中部宿舍卫生验评通报总结", &title_fmt)?;
    // Add logo to cell D1 (row 0, col 3) - 4th column
    let image = Image::new("assets/logo.png")?
        .set_scale_width(0.3)
        .set_scale_height(0.3);
    worksheet.insert_image(0, 3, &image)?;
    worksheet.merge_range(
        1,
        0,
        1,
        4,
        &format!("汇报人: {}", args.reporter),
        &left_align,
    )?;
    worksheet.merge_range(1, 5, 1, 7, "验评对象: 高一、高二、高三", &center_bold)?;
    worksheet.write_string_with_format(1, 8, format!("日期: {}", args.date), &center_bold)?;
    worksheet.write_string_with_format(2, 0, "验评部门", &center_bold)?;
    worksheet.merge_range(2, 1, 2, 8, "校办公室", &cell_fmt)?;
    worksheet.write_string_with_format(3, 0, "验评项目", &center_bold)?;
    worksheet.merge_range(3, 1, 3, 8, "高一高二高三男生宿舍卫生", &cell_fmt)?;
    worksheet.write_string_with_format(4, 0, "验评时间", &center_bold)?;
    worksheet.merge_range(4, 1, 4, 8, &args.time, &cell_fmt)?;
    worksheet.write_string_with_format(5, 0, "验评细则", &center_bold)?;
    worksheet.merge_range(5, 1, 5, 8, rules, &left_text)?;
    worksheet.set_row_height(5, 80)?;

    let headers1 = [
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
    for (i, h) in headers1.iter().enumerate() {
        worksheet.write_string_with_format(6, i as u16, *h, &header_fmt)?;
    }

    let mut row: u32 = 7;
    let mut apartments: Vec<u8> = processed_data
        .iter()
        .map(|r| r.apartment)
        .collect::<HashSet<_>>()
        .into_iter()
        .collect();
    apartments.sort();

    for apt in &apartments {
        let apt_name = format!("{}号公寓", if *apt == 1 { "一" } else { "二" });
        let mut groups: HashMap<(u8, String), Vec<&ProcessedRecord>> = HashMap::new();
        for r in processed_data.iter().filter(|r| r.apartment == *apt) {
            groups.entry((r.grade, r.dept.clone())).or_default().push(r);
        }

        let mut group_totals: Vec<((u8, String), i32)> = groups
            .iter()
            .map(|(k, v)| (k.clone(), v.iter().map(|r| r.deduction).sum()))
            .collect();
        group_totals.sort_by(|a, b| b.1.cmp(&a.1));

        let mut rank_map: HashMap<(u8, String), i32> = HashMap::new();
        let (mut cur_rank, mut last_score, mut cnt) = (1, i32::MAX, 0);
        for (i, (key, score)) in group_totals.iter().enumerate() {
            if i == 0 {
                cur_rank = 1;
                last_score = *score;
                cnt = 1;
            } else if *score == last_score {
                cnt += 1;
            } else {
                cur_rank += cnt;
                cnt = 1;
                last_score = *score;
            }
            rank_map.insert(key.clone(), cur_rank);
        }

        // Sort by grade and department: 高一A, 高一B, 高二A, 高二B, 高三A, 高三B
        let mut sorted_keys: Vec<_> = groups.keys().cloned().collect();
        sorted_keys.sort_by(|a, b| {
            if a.0 != b.0 {
                a.0.cmp(&b.0)
            } else {
                a.1.cmp(&b.1)
            }
        });
        let apt_start = row;

        for (grade, dept) in sorted_keys {
            let rs = groups.get(&(grade, dept.clone())).unwrap();
            let total = group_totals
                .iter()
                .find(|(k, _)| *k == (grade, dept.clone()))
                .unwrap()
                .1;
            let rank = *rank_map.get(&(grade, dept.clone())).unwrap();
            let grade_name = match grade {
                1 => "高一",
                2 => "高二",
                3 => "高三",
                _ => "",
            };
            let dept_display = format!("{}{}部", grade_name, dept);

            let mut sorted_rs: Vec<_> = rs.iter().collect();
            sorted_rs.sort_by_key(|r| r.dorm);
            let grp_start = row;

            for r in &sorted_rs {
                worksheet.write_string_with_format(row, 2, &r.teacher, &cell_fmt)?;
                worksheet.write_string_with_format(row, 3, &r.manager, &cell_fmt)?;
                worksheet.write_string_with_format(row, 4, format!("{}宿舍", r.dorm), &cell_fmt)?;
                worksheet.write_string_with_format(row, 5, &r.reason, &cell_fmt)?;
                worksheet.write_number_with_format(row, 6, r.deduction as f64, &cell_fmt)?;
                row += 1;
            }

            if row > grp_start {
                if row - grp_start > 1 {
                    worksheet.merge_range(grp_start, 1, row - 1, 1, &dept_display, &cell_fmt)?;
                    worksheet.merge_range(
                        grp_start,
                        7,
                        row - 1,
                        7,
                        &total.to_string(),
                        &cell_fmt,
                    )?;
                    worksheet.merge_range(
                        grp_start,
                        8,
                        row - 1,
                        8,
                        &rank.to_string(),
                        &cell_fmt,
                    )?;
                } else {
                    worksheet.write_string_with_format(grp_start, 1, &dept_display, &cell_fmt)?;
                    worksheet.write_number_with_format(grp_start, 7, total as f64, &cell_fmt)?;
                    worksheet.write_number_with_format(grp_start, 8, rank as f64, &cell_fmt)?;
                }
            }
        }

        if row > apt_start {
            if row - apt_start > 1 {
                worksheet.merge_range(apt_start, 0, row - 1, 0, &apt_name, &cell_fmt)?;
            } else {
                worksheet.write_string_with_format(apt_start, 0, &apt_name, &cell_fmt)?;
            }
        }
    }

    row += 2;

    // Table 2 - Add title and logo
    worksheet.merge_range(row, 0, row, 8, "高中部宿舍卫生验评通报总结", &title_fmt)?;
    let image2 = Image::new("assets/logo.png")?
        .set_scale_width(0.3)
        .set_scale_height(0.3);
    worksheet.insert_image(row, 3, &image2)?;
    row += 1;
    worksheet.merge_range(
        row,
        0,
        row,
        4,
        &format!("汇报人: {}", args.reporter),
        &left_align,
    )?;
    worksheet.merge_range(row, 5, row, 7, "验评对象: 高一、高二、高三", &center_bold)?;
    worksheet.write_string_with_format(row, 8, format!("日期: {}", args.date), &center_bold)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评部门", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, "校办公室", &cell_fmt)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评项目", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, "高一高二高三男生宿舍卫生", &cell_fmt)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评时间", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, &args.time, &cell_fmt)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评细则", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, rules, &left_text)?;
    worksheet.set_row_height(row, 80)?;
    row += 1;

    worksheet.write_string_with_format(row, 0, "公寓", &header_fmt)?;
    worksheet.write_string_with_format(row, 1, "宿舍管理员", &header_fmt)?;
    worksheet.write_string_with_format(row, 2, "宿舍号", &header_fmt)?;
    worksheet.merge_range(row, 3, row, 4, "扣分原因", &header_fmt)?;
    worksheet.write_string_with_format(row, 5, "扣分", &header_fmt)?;
    worksheet.merge_range(row, 6, row, 7, "总扣分", &header_fmt)?;
    worksheet.write_string_with_format(row, 8, "排名", &header_fmt)?;
    row += 1;

    let mut mgr_by_apt: HashMap<u8, HashSet<String>> = HashMap::new();
    for (apt, _, name) in &all_managers {
        mgr_by_apt.entry(*apt).or_default().insert(name.clone());
    }
    for r in &processed_data {
        mgr_by_apt
            .entry(r.apartment)
            .or_default()
            .insert(r.manager.clone());
    }

    let mut sorted_apts: Vec<u8> = mgr_by_apt.keys().cloned().collect();
    sorted_apts.sort();

    for apt in sorted_apts {
        let apt_name = format!("{}号公寓", if apt == 1 { "一" } else { "二" });
        let mgrs = mgr_by_apt.get(&apt).unwrap();

        let mut mgr_totals: Vec<(String, i32)> = mgrs
            .iter()
            .map(|m| {
                let t: i32 = processed_data
                    .iter()
                    .filter(|r| r.apartment == apt && &r.manager == m)
                    .map(|r| r.deduction)
                    .sum();
                (m.clone(), t)
            })
            .collect();
        mgr_totals.sort_by(|a, b| b.1.cmp(&a.1));

        let mut rank_map: HashMap<String, i32> = HashMap::new();
        let (mut cur_rank, mut last_val, mut cnt) = (1, i32::MAX, 0);
        for (i, (mgr, score)) in mgr_totals.iter().enumerate() {
            if i == 0 {
                cur_rank = 1;
                last_val = *score;
                cnt = 1;
            } else if *score == last_val {
                cnt += 1;
            } else {
                cur_rank += cnt;
                cnt = 1;
                last_val = *score;
            }
            rank_map.insert(mgr.clone(), cur_rank);
        }

        let mut mgr_floors: HashMap<String, u8> = HashMap::new();
        for (a, f, n) in &all_managers {
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
            let recs: Vec<_> = processed_data
                .iter()
                .filter(|r| r.apartment == apt && r.manager == mgr)
                .collect();
            let mgr_start = row;

            if recs.is_empty() {
                worksheet.write_string_with_format(row, 1, &mgr, &cell_fmt)?;
                worksheet.write_string_with_format(row, 2, "/", &cell_fmt)?;
                worksheet.merge_range(row, 3, row, 4, "/", &cell_fmt)?;
                worksheet.write_string_with_format(row, 5, "/", &cell_fmt)?;
                worksheet.merge_range(row, 6, row, 7, "/", &cell_fmt)?;
                worksheet.write_number_with_format(row, 8, rank as f64, &cell_fmt)?;
                row += 1;
            } else {
                let mut sorted_recs: Vec<_> = recs.iter().collect();
                sorted_recs.sort_by_key(|r| r.dorm);

                for r in &sorted_recs {
                    worksheet.write_string_with_format(
                        row,
                        2,
                        format!("{}宿舍", r.dorm),
                        &cell_fmt,
                    )?;
                    worksheet.merge_range(row, 3, row, 4, &r.reason, &cell_fmt)?;
                    worksheet.write_number_with_format(row, 5, r.deduction as f64, &cell_fmt)?;
                    row += 1;
                }

                if row > mgr_start {
                    if row - mgr_start > 1 {
                        worksheet.merge_range(mgr_start, 1, row - 1, 1, &mgr, &cell_fmt)?;
                        worksheet.merge_range(
                            mgr_start,
                            6,
                            row - 1,
                            7,
                            &total.to_string(),
                            &cell_fmt,
                        )?;
                        worksheet.merge_range(
                            mgr_start,
                            8,
                            row - 1,
                            8,
                            &rank.to_string(),
                            &cell_fmt,
                        )?;
                    } else {
                        worksheet.write_string_with_format(mgr_start, 1, &mgr, &cell_fmt)?;
                        worksheet.merge_range(
                            mgr_start,
                            6,
                            mgr_start,
                            7,
                            &total.to_string(),
                            &cell_fmt,
                        )?;
                        worksheet.write_number_with_format(mgr_start, 8, rank as f64, &cell_fmt)?;
                    }
                }
            }
        }

        if row > apt_start {
            if row - apt_start > 1 {
                worksheet.merge_range(apt_start, 0, row - 1, 0, &apt_name, &cell_fmt)?;
            } else {
                worksheet.write_string_with_format(apt_start, 0, &apt_name, &cell_fmt)?;
            }
        }
    }

    worksheet.set_column_width(0, 12)?;
    worksheet.set_column_width(1, 12)?;
    worksheet.set_column_width(2, 12)?;
    worksheet.set_column_width(3, 10)?;
    worksheet.set_column_width(4, 10)?;
    worksheet.set_column_width(5, 18)?;
    worksheet.set_column_width(6, 8)?;
    worksheet.set_column_width(7, 8)?;
    worksheet.set_column_width(8, 8)?;

    workbook.save("report.xlsx")?;
    println!("报告已生成: report.xlsx");
    Ok(())
}

fn load_test_data<P: AsRef<Path>>(path: P) -> Result<Vec<TestDataRecord>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut records = Vec::new();
    for result in rdr.deserialize() {
        records.push(result?);
    }
    Ok(records)
}

fn load_nianji_data<P: AsRef<Path>>(path: P) -> Result<HashMap<(u8, u8), (String, String)>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new()
        .has_headers(true)
        .flexible(true)
        .from_reader(file);
    let mut map = HashMap::new();
    for result in rdr.deserialize() {
        let r: NianjiRecord = result?;
        map.insert((r.grade, r.class), (r.dept.unwrap_or_default(), r.teacher));
    }
    Ok(map)
}

fn load_sushe_data<P: AsRef<Path>>(path: P) -> Result<HashMap<(u8, u8), String>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut map = HashMap::new();
    for result in rdr.deserialize() {
        let r: SusheRecord = result?;
        map.insert((r.apartment, r.floor), r.manager);
    }
    Ok(map)
}

fn get_all_managers<P: AsRef<Path>>(path: P) -> Result<Vec<(u8, u8, String)>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut list = Vec::new();
    for result in rdr.deserialize() {
        let r: SusheRecord = result?;
        list.push((r.apartment, r.floor, r.manager));
    }
    Ok(list)
}
