use anyhow::Result;
use clap::{CommandFactory, Parser, Subcommand};
use std::env;
use std::io;
use std::path::PathBuf;

#[derive(Debug, Parser)]
#[command(name = "excel-dr-core", version)]
struct Cli {
    #[command(subcommand)]
    command: Command,
    #[arg(long, global = true)]
    json: bool,
}

#[derive(Debug, Subcommand)]
enum Command {
    AnalyzeFile {
        path: PathBuf,
    },
    CleanFile {
        path: PathBuf,
        #[arg(short, long)]
        output: Option<PathBuf>,
    },
    AnalyzeFolder {
        path: PathBuf,
    },
    CleanFolder {
        path: PathBuf,
    },
}

fn main() -> Result<()> {
    if env::args_os().len() == 1 {
        Cli::command().print_help()?;
        println!();
        println!();
        println!("这是 Excel-Dr 的 Rust 命令行后端，不是双击式图形界面。");
        println!("常用示例:");
        println!("  excel-dr-core.exe analyze-file \"C:\\path\\file.xlsx\"");
        println!("  excel-dr-core.exe clean-file \"C:\\path\\file.xlsx\"");
        println!("  excel-dr-core.exe analyze-folder \"C:\\path\\folder\"");
        println!("  excel-dr-core.exe clean-folder \"C:\\path\\folder\"");
        println!();
        println!("双击查看到这里是正常的。图形界面请运行 dist\\Excel-Dr.exe。");
        println!("按 Enter 退出...");
        let mut line = String::new();
        let _ = io::stdin().read_line(&mut line);
        return Ok(());
    }

    let cli = Cli::parse();
    match cli.command {
        Command::AnalyzeFile { path } => {
            let report = excel_dr_core::analyze_file(path)?;
            print_value(&report, cli.json)?;
        }
        Command::CleanFile { path, output } => {
            let output = output.unwrap_or_else(|| excel_dr_core::default_output_path(&path));
            let report = excel_dr_core::clean_file(path, output)?;
            print_value(&report, cli.json)?;
        }
        Command::AnalyzeFolder { path } => {
            let result = excel_dr_core::analyze_folder(path)?;
            print_value(&result, cli.json)?;
        }
        Command::CleanFolder { path } => {
            let result = excel_dr_core::clean_folder(path)?;
            print_value(&result, cli.json)?;
        }
    }
    Ok(())
}

fn print_value<T>(value: &T, json: bool) -> Result<()>
where
    T: serde::Serialize + Renderable,
{
    if json {
        println!("{}", serde_json::to_string_pretty(value)?);
    } else {
        println!("{}", value.render_text());
    }
    Ok(())
}

trait Renderable {
    fn render_text(&self) -> String;
}

impl Renderable for excel_dr_core::WorkbookReport {
    fn render_text(&self) -> String {
        self.render()
    }
}

impl Renderable for excel_dr_core::BatchResult {
    fn render_text(&self) -> String {
        self.render()
    }
}
