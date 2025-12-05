use serde::Deserialize;

#[derive(Debug, Deserialize)]
pub struct ReportDataRecord {
    #[serde(rename = "年级")]
    pub grade: u8,
    #[serde(rename = "班级")]
    pub class: u8,
    #[serde(rename = "公寓")]
    pub apartment: u8,
    #[serde(rename = "宿舍")]
    pub dorm: u16,
    #[serde(rename = "原因")]
    pub reason: String,
}

#[derive(Debug, Deserialize)]
pub struct GradeRecord {
    #[serde(rename = "年级")]
    pub grade: u8,
    #[serde(rename = "级部")]
    pub dept: Option<String>,
    #[serde(rename = "班级")]
    pub class: u8,
    #[serde(rename = "班主任")]
    pub teacher: String,
}

#[derive(Debug, Deserialize)]
pub struct ApartmentRecord {
    #[serde(rename = "公寓")]
    pub apartment: u8,
    #[serde(rename = "楼层")]
    pub floor: u8,
    #[serde(rename = "宿管")]
    pub manager: String,
}

#[derive(Debug, Deserialize)]
pub struct DepartmentRecord {
    #[serde(rename = "年级")]
    pub grade: u8,
    #[serde(rename = "级部")]
    pub dept: String,
    #[serde(rename = "主任")]
    pub leader: String,
    #[serde(rename = "公寓")]
    pub apartment: u8,
}

pub struct ProcessedRecord {
    pub apartment: u8,
    pub grade: u8,
    pub class: u8,
    pub dept: String,
    pub teacher: String,
    pub manager: String,
    pub dorm: u16,
    pub reason: String,
    pub deduction: i32,
}
