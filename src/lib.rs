use std::{fmt, fs::File, io::BufReader};

use calamine::{open_workbook, Data, Reader, Xlsx};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

#[derive(Clone)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
enum CellDiffKind {
    Value,
    Formula,
}

impl fmt::Display for CellDiffKind {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            CellDiffKind::Formula => write!(f, "formula"),
            CellDiffKind::Value => write!(f, "value"),
        }
    }
}

/// main struct
#[derive(Clone)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct Diff {
    old_filepath: String,
    new_filepath: String,
    sheet_diff: Vec<SheetDiff>,
    cell_diffs: Vec<SheetCellDiff>,
}

#[derive(Clone)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
struct SheetDiff {
    old: Option<String>,
    new: Option<String>,
}

#[derive(Clone)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
struct SheetCellDiff {
    kind: CellDiffKind,
    sheet: String,
    cells: Vec<CellDiff>,
}

#[derive(Clone)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
struct CellDiff {
    row: u32,
    col: u32,
    old: Option<String>,
    new: Option<String>,
}

impl Diff {
    /// init
    pub fn new(old_filepath: &str, new_filepath: &str) -> Self {
        let mut ret = Diff {
            old_filepath: old_filepath.to_owned(),
            new_filepath: new_filepath.to_owned(),
            sheet_diff: vec![],
            cell_diffs: vec![],
        };
        ret.collect_diff();
        ret
    }

    /// get serde-ready diff
    /// #[cfg(feature = "serde")]
    pub fn diff(&mut self) -> Diff {
        self.clone()
    }

    /// get unified diff str
    pub fn unified_diff(&mut self) -> String {
        let mut ret: Vec<String> = vec![];

        if !self.sheet_diff.is_empty() {
            ret.push(format!("--- {} (sheet names)", self.old_filepath));
            ret.push(format!("+++ {} (sheet names)", self.new_filepath));
            self.sheet_diff.iter().for_each(|x| {
                if let Some(sheet) = x.old.as_ref() {
                    ret.push(format!("- {}", sheet));
                }
                if let Some(sheet) = x.new.as_ref() {
                    ret.push(format!("+ {}", sheet));
                }
            });
        }

        self.cell_diffs.iter().for_each(|x| {
            ret.push(format!(
                "--- {} [{}] ({})",
                self.old_filepath, x.sheet, x.kind
            ));
            ret.push(format!(
                "+++ {} [{}] ({})",
                self.new_filepath, x.sheet, x.kind
            ));
            x.cells.iter().for_each(|x| {
                ret.push(format!("@@ ({}, {}) @@", x.row, x.col));
                if let Some(sheet) = x.old.as_ref() {
                    ret.push(format!("- {}", sheet));
                }
                if let Some(sheet) = x.new.as_ref() {
                    ret.push(format!("+ {}", sheet));
                }
            });
        });

        ret.join("\n")
    }

    /// collect sheet diff and cell range diff
    fn collect_diff(&mut self) {
        let mut old_workbook: Xlsx<BufReader<File>> = open_workbook(self.old_filepath.as_str())
            .expect(format!("Cannot open {}", self.old_filepath.as_str()).as_str());
        let mut new_workbook: Xlsx<BufReader<File>> = open_workbook(self.new_filepath.as_str())
            .expect(format!("Cannot open {}", self.new_filepath.as_str()).as_str());

        let old_sheets = old_workbook.sheet_names().to_owned();
        let new_sheets = new_workbook.sheet_names().to_owned();

        self.collect_sheet_diff(&old_sheets, &new_sheets);

        let same_name_sheets = filter_same_name_sheets(&old_sheets, &new_sheets);
        self.collect_cell_value_diff(&mut old_workbook, &mut new_workbook, &same_name_sheets);
        self.collect_cell_formula_diff(&mut old_workbook, &mut new_workbook, &same_name_sheets);
    }

    /// collect sheet diff by name
    fn collect_sheet_diff(&mut self, old_sheets: &Vec<String>, new_sheets: &Vec<String>) {
        if *old_sheets == *new_sheets {
            return;
        }

        for sheet in old_sheets {
            if !new_sheets.contains(sheet) {
                self.sheet_diff.push(SheetDiff {
                    old: Some(sheet.to_owned()),
                    new: None,
                });
            }
        }
        for sheet in new_sheets {
            if !old_sheets.contains(sheet) {
                self.sheet_diff.push(SheetDiff {
                    old: None,
                    new: Some(sheet.to_owned()),
                });
            }
        }
    }

    /// collect value diff in cell range
    fn collect_cell_value_diff(
        &mut self,
        old_workbook: &mut Xlsx<BufReader<File>>,
        new_workbook: &mut Xlsx<BufReader<File>>,
        same_name_sheets: &Vec<String>,
    ) {
        for sheet in same_name_sheets {
            if let (Ok(range1), Ok(range2)) = (
                old_workbook.worksheet_range(sheet),
                new_workbook.worksheet_range(sheet),
            ) {
                let mut cell_diffs: Vec<CellDiff> = vec![];

                let max_rows = range1.height().max(range2.height());
                let max_cols = range1.width().max(range2.width());

                for row in 0..max_rows {
                    for col in 0..max_cols {
                        let cell1 = range1
                            .get_value((row as u32, col as u32))
                            .unwrap_or(&Data::Empty);
                        let cell2 = range2
                            .get_value((row as u32, col as u32))
                            .unwrap_or(&Data::Empty);

                        if cell1 != cell2 {
                            cell_diffs.push(CellDiff {
                                row: (row + 1) as u32,
                                col: (col + 1) as u32,
                                old: if cell1 != &Data::Empty {
                                    Some(cell1.to_string())
                                } else {
                                    None
                                },
                                new: if cell2 != &Data::Empty {
                                    Some(cell2.to_string())
                                } else {
                                    None
                                },
                            });
                        }
                    }
                }

                if !cell_diffs.is_empty() {
                    let sheet_cell_diff = SheetCellDiff {
                        sheet: sheet.to_owned(),
                        kind: CellDiffKind::Value,
                        cells: cell_diffs,
                    };
                    self.cell_diffs.push(sheet_cell_diff);
                }
            } else {
                println!("Failed to read sheet: {}", sheet);
            }
        }
    }

    /// collect formula diff in cell range
    fn collect_cell_formula_diff(
        &mut self,
        old_workbook: &mut Xlsx<BufReader<File>>,
        new_workbook: &mut Xlsx<BufReader<File>>,
        same_name_sheets: &Vec<String>,
    ) {
        for sheet in same_name_sheets {
            if let (Ok(range1), Ok(range2)) = (
                old_workbook.worksheet_formula(sheet),
                new_workbook.worksheet_formula(sheet),
            ) {
                let mut cell_diffs: Vec<CellDiff> = vec![];

                let (range1_start_row, range1_start_col) = match range1.start() {
                    Some((row, col)) => (row, col),
                    None => (u32::MAX, u32::MAX),
                };
                let (range2_start_row, range2_start_col) = match range2.start() {
                    Some((row, col)) => (row, col),
                    None => (u32::MAX, u32::MAX),
                };
                let (range1_end_row, range1_end_col) = match range1.end() {
                    Some((row, col)) => (row, col),
                    None => (u32::MIN, u32::MIN),
                };
                let (range2_end_row, range2_end_col) = match range2.end() {
                    Some((row, col)) => (row, col),
                    None => (u32::MIN, u32::MIN),
                };
                let start_row = range1_start_row.min(range2_start_row);
                let start_col = range1_start_col.min(range2_start_col);
                let end_row = range1_end_row.max(range2_end_row);
                let end_col = range1_end_col.max(range2_end_col);

                for row in start_row..(end_row + 1) {
                    for col in start_col..(end_col + 1) {
                        let cell1 = match range1.get_value((row as u32, col as u32)) {
                            Some(x) => &Data::String(x.to_string()),
                            None => &Data::Empty,
                        };
                        let cell2 = match range2.get_value((row as u32, col as u32)) {
                            Some(x) => &Data::String(x.to_string()),
                            None => &Data::Empty,
                        };

                        if cell1 != cell2 {
                            cell_diffs.push(CellDiff {
                                row: (row + 1) as u32,
                                col: (col + 1) as u32,
                                old: if cell1 != &Data::Empty {
                                    Some(cell1.to_string())
                                } else {
                                    None
                                },
                                new: if cell2 != &Data::Empty {
                                    Some(cell2.to_string())
                                } else {
                                    None
                                },
                            });
                        }
                    }
                }

                if !cell_diffs.is_empty() {
                    let sheet_cell_diff = SheetCellDiff {
                        sheet: sheet.to_owned(),
                        kind: CellDiffKind::Formula,
                        cells: cell_diffs,
                    };
                    self.cell_diffs.push(sheet_cell_diff);
                }
            } else {
                println!("Failed to read sheet: {}", sheet);
            }
        }
    }
}

/// filter sheets whose name is equal
fn filter_same_name_sheets<'a>(
    old_sheets: &'a Vec<String>,
    new_sheets: &'a Vec<String>,
) -> Vec<String> {
    old_sheets
        .iter()
        .filter(|s| new_sheets.contains(s))
        .map(|s| s.to_owned())
        .collect()
}
