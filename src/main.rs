use anyhow::Result;
use clap::{Parser, Subcommand};
use std::path::PathBuf;

mod init;
mod model;
mod report;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand, Debug)]
enum Commands {
    /// 初始化一个新的CSV文件模板
    Init {
        /// CSV文件名
        filename: String,
    },
    /// 生成卫生验评报告
    Report {
        /// 输入CSV文件路径
        input: PathBuf,

        /// 输出Excel文件路径（可选，默认与输入文件同名但扩展名为.xlsx）
        #[arg(short, long)]
        output: Option<PathBuf>,

        #[arg(short, long, default_value = "杨超超、申淑玲、赵冰、徐雪冰")]
        reporter: String,

        #[arg(short, long, default_value = "12月5日")]
        date: String,

        #[arg(short, long, default_value = "下午: 15:05-xx:xx")]
        time: String,
    },
}

fn main() -> Result<()> {
    let args = Args::parse();

    match args.command {
        Commands::Init { filename } => {
            init::init_csv(&filename)?;
        }
        Commands::Report {
            input,
            output,
            reporter,
            date,
            time,
        } => {
            report::generate_report(input, output, reporter, date, time)?;
        }
    }

    Ok(())
}
