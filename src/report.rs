use anyhow::Result;
use csv::ReaderBuilder;
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Image, Workbook};
use std::{
    collections::{HashMap, HashSet},
    fs::File,
    path::{Path, PathBuf},
};

use crate::model::{
    ApartmentRecord, DepartmentRecord, GradeRecord, ProcessedRecord, ReportDataRecord,
};

pub fn generate_report(
    input: PathBuf,
    output: Option<PathBuf>,
    reporter: String,
    date: String,
    time: String,
) -> Result<()> {
    // 确定输出文件路径
    let output_path = match &output {
        Some(path) => path.clone(),
        None => {
            // 如果未指定输出文件，使用输入文件名但改为.xlsx扩展名
            let mut out = input.clone();
            out.set_extension("xlsx");
            out
        }
    };

    let records = load_report_data(&input)?;
    let nianji_map = load_grade_data("assets/nianji.csv")?;
    let sushe_map = load_apt_data("assets/sushe.csv")?;
    let all_managers = get_all_managers("assets/sushe.csv")?;
    let jibu_map = load_dept_data("assets/jibu.csv")?;

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
            class: r.class,
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
    worksheet.merge_range(1, 0, 1, 4, &format!("汇报人: {}", reporter), &left_align)?;
    worksheet.merge_range(1, 5, 1, 7, "验评对象: 高一、高二、高三", &center_bold)?;
    worksheet.write_string_with_format(1, 8, format!("日期: {}", date), &center_bold)?;
    worksheet.write_string_with_format(2, 0, "验评部门", &center_bold)?;
    worksheet.merge_range(2, 1, 2, 8, "校办公室", &cell_fmt)?;
    worksheet.write_string_with_format(3, 0, "验评项目", &center_bold)?;
    worksheet.merge_range(3, 1, 3, 8, "高一高二高三男生宿舍卫生", &cell_fmt)?;
    worksheet.write_string_with_format(4, 0, "验评时间", &center_bold)?;
    worksheet.merge_range(4, 1, 4, 8, &time, &cell_fmt)?;
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
    apartments.sort_by(|a, b| b.cmp(a)); // 二号公寓在前

    // Calculate global rankings for all departments (not per apartment)
    let mut all_dept_groups: HashMap<(u8, String), Vec<&ProcessedRecord>> = HashMap::new();

    // Initialize all departments from jibu_map
    for (grade, dept) in jibu_map.keys() {
        all_dept_groups.entry((*grade, dept.clone())).or_default();
    }

    // Group all records by department (across all apartments)
    for r in &processed_data {
        if !r.dept.is_empty() {
            all_dept_groups
                .entry((r.grade, r.dept.clone()))
                .or_default()
                .push(r);
        }
    }

    // Calculate totals for all departments
    let mut all_dept_totals: Vec<((u8, String), i32)> = all_dept_groups
        .iter()
        .map(|(k, v)| (k.clone(), v.iter().map(|r| r.deduction).sum()))
        .collect();
    all_dept_totals.sort_by(|a, b| b.1.cmp(&a.1)); // 扣分少的排名靠前（0 > -1 > -2）

    // Compute global ranks for all departments (并列不跳号: 1,1,2,3)
    let mut global_dept_rank_map: HashMap<(u8, String), i32> = HashMap::new();
    if !all_dept_totals.is_empty() {
        let mut cur_rank = 1;
        let mut prev_score = all_dept_totals[0].1;
        global_dept_rank_map.insert(all_dept_totals[0].0.clone(), cur_rank);

        for (key, score) in all_dept_totals.iter().skip(1) {
            if *score != prev_score {
                cur_rank += 1;
                prev_score = *score;
            }
            global_dept_rank_map.insert(key.clone(), cur_rank);
        }
    }

    // Check which apartments have records for 高二A部 (special case)
    let mut apt2a_has_records: HashMap<u8, bool> = HashMap::new();
    for r in &processed_data {
        if r.grade == 2 && r.dept == "A" {
            apt2a_has_records.insert(r.apartment, true);
        }
    }
    let apt2a_in_both = apt2a_has_records.contains_key(&1) && apt2a_has_records.contains_key(&2);
    let apt2a_in_apt1_only =
        apt2a_has_records.contains_key(&1) && !apt2a_has_records.contains_key(&2);
    let apt2a_in_apt2_only =
        apt2a_has_records.contains_key(&2) && !apt2a_has_records.contains_key(&1);
    let apt2a_in_neither = apt2a_has_records.is_empty();

    // Track row ranges for 高二A部 to merge across apartments
    let mut apt2a_start_row: Option<u32> = None;
    let mut apt2a_end_row: Option<u32> = None;

    for apt in &apartments {
        let apt_name = format!("{}号公寓", if *apt == 1 { "一" } else { "二" });

        // Separate records: those with dept go into dept_groups, those without go into class_groups
        let mut dept_groups: HashMap<(u8, String), Vec<&ProcessedRecord>> = HashMap::new();
        let mut class_groups: HashMap<u8, Vec<&ProcessedRecord>> = HashMap::new();

        // Initialize departments that belong to this apartment (based on jibu.csv)
        for ((grade, dept), (_, default_apt)) in jibu_map.iter() {
            // Special handling for 高二A部
            if *grade == 2 && dept == "A" {
                if apt2a_in_both {
                    // Show in both apartments
                    dept_groups.entry((*grade, dept.clone())).or_default();
                } else if apt2a_in_apt1_only && *apt == 1 {
                    // Only show in apt 1
                    dept_groups.entry((*grade, dept.clone())).or_default();
                } else if apt2a_in_apt2_only && *apt == 2 {
                    // Only show in apt 2
                    dept_groups.entry((*grade, dept.clone())).or_default();
                } else if apt2a_in_neither && *apt == 1 {
                    // No records anywhere, show in default apt (1)
                    dept_groups.entry((*grade, dept.clone())).or_default();
                }
            } else if *default_apt == *apt {
                dept_groups.entry((*grade, dept.clone())).or_default();
            }
        }

        // Add records to departments (only records from this apartment)
        for r in processed_data.iter().filter(|r| r.apartment == *apt) {
            if r.dept.is_empty() {
                // No department - group by class number (17班, 18班)
                class_groups.entry(r.class).or_default().push(r);
            } else {
                // Has department - group by (grade, dept)
                dept_groups
                    .entry((r.grade, r.dept.clone()))
                    .or_default()
                    .push(r);
            }
        }

        // Calculate totals for dept groups
        let mut dept_totals: Vec<((u8, String), i32)> = dept_groups
            .iter()
            .map(|(k, v)| (k.clone(), v.iter().map(|r| r.deduction).sum()))
            .collect();
        dept_totals.sort_by(|a, b| b.1.cmp(&a.1)); // 扣分少的排名靠前（0 > -1 > -2）

        // Calculate totals for class groups (no dept)
        let mut class_totals: Vec<(u8, i32)> = class_groups
            .iter()
            .map(|(k, v)| (*k, v.iter().map(|r| r.deduction).sum()))
            .collect();
        class_totals.sort_by(|a, b| b.1.cmp(&a.1)); // 扣分少的排名靠前（0 > -1 > -2）

        // Compute ranks for class groups (no dept) (并列不跳号)
        let mut class_rank_map: HashMap<u8, i32> = HashMap::new();
        if !class_totals.is_empty() {
            let mut cur_rank = 1;
            let mut prev_score = class_totals[0].1;
            class_rank_map.insert(class_totals[0].0, cur_rank);

            for (class, score) in class_totals.iter().skip(1) {
                if *score != prev_score {
                    cur_rank += 1; // 不跳号，只加1
                    prev_score = *score;
                }
                class_rank_map.insert(*class, cur_rank);
            }
        }

        // Sort dept keys: 高一A, 高一B, 高二A, 高二B, 高三A, 高三B
        let mut sorted_dept_keys: Vec<_> = dept_groups.keys().cloned().collect();
        sorted_dept_keys.sort_by(|a, b| {
            if a.0 != b.0 {
                a.0.cmp(&b.0)
            } else {
                a.1.cmp(&b.1)
            }
        });

        // Sort class keys (17, 18)
        let mut sorted_class_keys: Vec<_> = class_groups.keys().cloned().collect();
        sorted_class_keys.sort();

        let apt_start = row;

        // First render dept groups (all departments that belong to this apartment)
        for (grade, dept) in sorted_dept_keys {
            let rs = dept_groups.get(&(grade, dept.clone())).unwrap();

            let grade_name = match grade {
                1 => "高一",
                2 => "高二",
                3 => "高三",
                _ => "",
            };
            let leader = jibu_map
                .get(&(grade, dept.clone()))
                .map(|(l, _)| l.clone())
                .unwrap_or_default();
            let dept_display = format!("{}{}部\n({})", grade_name, dept, leader);

            let grp_start = row;

            // Track 高二A部 row range for cross-apartment merging
            let is_2a = grade == 2 && dept == "A";
            if is_2a && apt2a_in_both
                && apt2a_start_row.is_none() {
                    apt2a_start_row = Some(row);
                }

            let total = dept_totals
                .iter()
                .find(|(k, _)| *k == (grade, dept.clone()))
                .unwrap()
                .1;
            let rank = *global_dept_rank_map.get(&(grade, dept.clone())).unwrap();

            if rs.is_empty() {
                // No issues for this department - display a row with "/"
                worksheet.write_string_with_format(row, 1, &dept_display, &cell_fmt)?;
                worksheet.write_string_with_format(row, 2, "/", &cell_fmt)?;
                worksheet.write_string_with_format(row, 3, "/", &cell_fmt)?;
                worksheet.write_string_with_format(row, 4, "/", &cell_fmt)?;
                worksheet.write_string_with_format(row, 5, "/", &cell_fmt)?;
                worksheet.write_string_with_format(row, 6, "/", &cell_fmt)?;
                worksheet.write_string_with_format(row, 7, "/", &cell_fmt)?;
                worksheet.write_number_with_format(row, 8, rank as f64, &cell_fmt)?;
                row += 1;
            } else {
                let mut sorted_rs: Vec<_> = rs.iter().collect();
                sorted_rs.sort_by_key(|r| r.dorm);

                // Write each dorm as a separate row
                for (idx, r) in sorted_rs.iter().enumerate() {
                    let current_row = grp_start + idx as u32;

                    // Write teacher (col 2)
                    worksheet.write_string_with_format(current_row, 2, &r.teacher, &cell_fmt)?;
                    // Write manager (col 3)
                    worksheet.write_string_with_format(current_row, 3, &r.manager, &cell_fmt)?;
                    // Write dorm (col 4)
                    worksheet.write_string_with_format(
                        current_row,
                        4,
                        format!("{}宿舍", r.dorm),
                        &cell_fmt,
                    )?;
                    // Write reason (col 5)
                    worksheet.write_string_with_format(current_row, 5, &r.reason, &cell_fmt)?;
                    // Write deduction (col 6)
                    worksheet.write_number_with_format(
                        current_row,
                        6,
                        r.deduction as f64,
                        &cell_fmt,
                    )?;
                }

                row += sorted_rs.len() as u32;

                // Track end row for 高二A部
                if is_2a && apt2a_in_both {
                    apt2a_end_row = Some(row - 1);
                }

                // For 高二A部 in both apartments, skip merging here (will merge across apartments later)
                if !(is_2a && apt2a_in_both) {
                    // Merge dept name (col 1), total (col 7), and rank (col 8) across all rows
                    if sorted_rs.len() > 1 {
                        worksheet.merge_range(
                            grp_start,
                            1,
                            row - 1,
                            1,
                            &dept_display,
                            &cell_fmt,
                        )?;
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
                        worksheet.write_string_with_format(
                            grp_start,
                            1,
                            &dept_display,
                            &cell_fmt,
                        )?;
                        worksheet.write_number_with_format(
                            grp_start,
                            7,
                            total as f64,
                            &cell_fmt,
                        )?;
                        worksheet.write_number_with_format(grp_start, 8, rank as f64, &cell_fmt)?;
                    }
                }
            }
        }

        // Then render class groups (no dept) - 17班, 18班 at the bottom
        // Only show classes if they have issues
        for class_num in sorted_class_keys {
            let rs = class_groups.get(&class_num).unwrap();

            // Skip classes without issues (高三17、18班若没问题可不显示)
            if rs.is_empty() {
                continue;
            } else {
                let total = class_totals
                    .iter()
                    .find(|(k, _)| *k == class_num)
                    .unwrap()
                    .1;
                let rank = *class_rank_map.get(&class_num).unwrap();
                let class_display = format!("{}班", class_num);

                let mut sorted_rs: Vec<_> = rs.iter().collect();
                sorted_rs.sort_by_key(|r| r.dorm);
                let grp_start = row;

                // Write each dorm as a separate row
                for (idx, r) in sorted_rs.iter().enumerate() {
                    let current_row = grp_start + idx as u32;

                    // Write teacher (col 2)
                    worksheet.write_string_with_format(current_row, 2, &r.teacher, &cell_fmt)?;
                    // Write manager (col 3)
                    worksheet.write_string_with_format(current_row, 3, &r.manager, &cell_fmt)?;
                    // Write dorm (col 4)
                    worksheet.write_string_with_format(
                        current_row,
                        4,
                        format!("{}宿舍", r.dorm),
                        &cell_fmt,
                    )?;
                    // Write reason (col 5)
                    worksheet.write_string_with_format(current_row, 5, &r.reason, &cell_fmt)?;
                    // Write deduction (col 6)
                    worksheet.write_number_with_format(
                        current_row,
                        6,
                        r.deduction as f64,
                        &cell_fmt,
                    )?;
                }

                row += sorted_rs.len() as u32;

                // Merge class name (col 1), total (col 7), and rank (col 8) across all rows
                if sorted_rs.len() > 1 {
                    worksheet.merge_range(grp_start, 1, row - 1, 1, &class_display, &cell_fmt)?;
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
                    worksheet.write_string_with_format(grp_start, 1, &class_display, &cell_fmt)?;
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

    // If 高二A部 appears in both apartments, merge cells across apartments
    if apt2a_in_both
        && let (Some(start), Some(end)) = (apt2a_start_row, apt2a_end_row) {
            // Get 高二A部 info
            let leader = jibu_map
                .get(&(2, "A".to_string()))
                .map(|(l, _)| l.clone())
                .unwrap_or_default();
            let dept_display = format!("高二A部\n({})", leader);

            // Get total and rank for 高二A部
            let total: i32 = all_dept_groups
                .get(&(2, "A".to_string()))
                .map(|v| v.iter().map(|r| r.deduction).sum())
                .unwrap_or(0);
            let rank = *global_dept_rank_map.get(&(2, "A".to_string())).unwrap();

            // Merge dept name (col 1), total (col 7), and rank (col 8) across both apartments
            worksheet.merge_range(start, 1, end, 1, &dept_display, &cell_fmt)?;
            worksheet.merge_range(start, 7, end, 7, &total.to_string(), &cell_fmt)?;
            worksheet.merge_range(start, 8, end, 8, &rank.to_string(), &cell_fmt)?;
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
        &format!("汇报人: {}", reporter),
        &left_align,
    )?;
    worksheet.merge_range(row, 5, row, 7, "验评对象: 高一、高二、高三", &center_bold)?;
    worksheet.write_string_with_format(row, 8, format!("日期: {}", date), &center_bold)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评部门", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, "校办公室", &cell_fmt)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评项目", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, "高一高二高三男生宿舍卫生", &cell_fmt)?;
    row += 1;
    worksheet.write_string_with_format(row, 0, "验评时间", &center_bold)?;
    worksheet.merge_range(row, 1, row, 8, &time, &cell_fmt)?;
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
        mgr_totals.sort_by(|a, b| b.1.cmp(&a.1)); // 扣分少的排名靠前（0 > -1 > -2）

        let mut rank_map: HashMap<String, i32> = HashMap::new();
        if !mgr_totals.is_empty() {
            let mut cur_rank = 1;
            let mut prev_score = mgr_totals[0].1;
            rank_map.insert(mgr_totals[0].0.clone(), cur_rank);

            for (mgr, score) in mgr_totals.iter().skip(1) {
                if *score != prev_score {
                    cur_rank += 1; // 不跳号，只加1
                    prev_score = *score;
                }
                rank_map.insert(mgr.clone(), cur_rank);
            }
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

    workbook.save(&output_path)?;
    println!("报告已生成: {}", output_path.display());
    Ok(())
}

fn load_report_data<P: AsRef<Path>>(path: P) -> Result<Vec<ReportDataRecord>> {
    let file = File::open(path)?;
    let mut rdr = ReaderBuilder::new().has_headers(true).from_reader(file);
    let mut records = Vec::new();
    for result in rdr.deserialize() {
        records.push(result?);
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
