use anyhow::Result;
use csv::Writer;

pub fn init_csv(filename: &str) -> Result<()> {
    let csv_filename = if filename.ends_with(".csv") {
        filename.to_string()
    } else {
        format!("{}.csv", filename)
    };

    let mut wtr = Writer::from_path(&csv_filename)?;
    wtr.write_record(["年级", "班级", "公寓", "宿舍", "原因"])?;
    wtr.flush()?;
    println!("已创建CSV文件: {}", csv_filename);
    Ok(())
}
