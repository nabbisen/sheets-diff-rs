use std::{env, fs};

use sheets_diff::core::{diff::Diff, unified_format::unified_diff};

fn main() {
    let args: Vec<String> = env::args().collect();
    let (old_filepath, new_filepath) = filepaths(args.as_ref());

    let diff = Diff::new(old_filepath, new_filepath);
    println!("{}", unified_diff(&diff));
}

fn filepaths<'a>(args: &'a Vec<String>) -> (&'a str, &'a str) {
    if args.len() != 3 {
        eprintln!("Usage: {} <file1> <file2>", args[0]);
        std::process::exit(1);
    }

    let old_filepath = &args[1];
    let new_filepath = &args[2];

    if !is_valid_filepath(old_filepath) || !is_valid_filepath(new_filepath) {
        eprintln!("Invalid file path(s) are found.");
        std::process::exit(1);
    }

    (old_filepath, new_filepath)
}

fn is_valid_filepath(filepath: &str) -> bool {
    fs::metadata(filepath).is_ok()
}
