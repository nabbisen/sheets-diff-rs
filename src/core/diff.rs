use std::{fmt, fs::File, io::BufReader};

use calamine::{open_workbook, Data, Reader, Xlsx};
#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub enum CellDiffKind {
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
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct Diff {
    pub old_filepath: String,
    pub new_filepath: String,
    pub sheet_diff: Vec<SheetDiff>,
    pub cell_diffs: Vec<SheetCellDiff>,
}

#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct SheetDiff {
    pub old: Option<String>,
    pub new: Option<String>,
}

#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct SheetCellDiff {
    pub sheet: String,
    pub cells: Vec<CellDiff>,
}

#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct CellDiff {
    pub row: usize,
    pub col: usize,
    pub addr: String,
    pub kind: CellDiffKind,
    pub old: Option<String>,
    pub new: Option<String>,
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

        ret.cell_diffs.sort_by(|a, b| a.sheet.cmp(&b.sheet));

        let mut merged_cell_diffs: Vec<SheetCellDiff> = vec![];
        ret.cell_diffs.iter().for_each(|a| {
            let found = merged_cell_diffs.iter_mut().find(|b| b.sheet == a.sheet);
            if let Some(found) = found {
                found.cells.extend(a.cells.clone());
            } else {
                merged_cell_diffs.push(a.clone());
            }
        });
        merged_cell_diffs.iter_mut().for_each(|x| {
            x.cells
                .sort_by(|a, b| a.addr.cmp(&b.addr).then_with(|| a.kind.cmp(&b.kind)));
        });
        ret.cell_diffs = merged_cell_diffs;

        ret
    }

    /// get serde-ready diff
    /// #[cfg(feature = "serde")]
    pub fn diff(&mut self) -> Diff {
        self.clone()
    }

    /// get unified diff str
    pub fn unified_diff(&mut self) -> UnifiedDiff {
        let mut ret: Vec<UnifiedDiffLine> = vec![];

        if !self.sheet_diff.is_empty() {
            ret.push(UnifiedDiffLine {
                kind: UnifiedDiffKind::OldTitle,
                text: format!("{} (sheet names)", self.old_filepath),
            });
            ret.push(UnifiedDiffLine {
                kind: UnifiedDiffKind::NewTitle,
                text: format!("{} (sheet names)", self.new_filepath),
            });

            self.sheet_diff.iter().for_each(|x| {
                if let Some(sheet) = x.old.as_ref() {
                    ret.push(UnifiedDiffLine {
                        kind: UnifiedDiffKind::OldContent,
                        text: sheet.to_owned(),
                    });
                }
                if let Some(sheet) = x.new.as_ref() {
                    ret.push(UnifiedDiffLine {
                        kind: UnifiedDiffKind::NewContent,
                        text: sheet.to_owned(),
                    });
                }
            });
        }

        self.cell_diffs.iter().for_each(|x| {
            ret.push(UnifiedDiffLine {
                kind: UnifiedDiffKind::OldTitle,
                text: format!("{} [{}]", self.old_filepath, x.sheet),
            });
            ret.push(UnifiedDiffLine {
                kind: UnifiedDiffKind::NewTitle,
                text: format!("{} [{}]", self.new_filepath, x.sheet),
            });

            x.cells.iter().for_each(|x| {
                ret.push(UnifiedDiffLine {
                    kind: UnifiedDiffKind::DiffPos,
                    text: format!("{}({},{}) {}", x.addr, x.row, x.col, x.kind),
                });

                if let Some(sheet) = x.old.as_ref() {
                    ret.push(UnifiedDiffLine {
                        kind: UnifiedDiffKind::OldContent,
                        text: sheet.to_owned(),
                    });
                }
                if let Some(sheet) = x.new.as_ref() {
                    ret.push(UnifiedDiffLine {
                        kind: UnifiedDiffKind::NewContent,
                        text: sheet.to_owned(),
                    });
                }
            });
        });

        UnifiedDiff { lines: ret }
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
            if let (Ok(old_range), Ok(new_range)) = (
                old_workbook.worksheet_range(sheet),
                new_workbook.worksheet_range(sheet),
            ) {
                let mut cell_diffs: Vec<CellDiff> = vec![];

                let (start_row, start_col, end_row, end_col) = diff_range(
                    old_range.start(),
                    new_range.start(),
                    old_range.end(),
                    new_range.end(),
                );

                for row in start_row..end_row {
                    for col in start_col..end_col {
                        let old_cell = old_range.get_value((row, col)).unwrap_or(&Data::Empty);
                        let new_cell = new_range.get_value((row, col)).unwrap_or(&Data::Empty);

                        if old_cell != new_cell {
                            let row = (row + 1) as usize;
                            let col = (col + 1) as usize;
                            cell_diffs.push(CellDiff {
                                row,
                                col,
                                addr: cell_pos_to_address(row, col),
                                kind: CellDiffKind::Value,
                                old: if old_cell != &Data::Empty {
                                    Some(old_cell.to_string())
                                } else {
                                    None
                                },
                                new: if new_cell != &Data::Empty {
                                    Some(new_cell.to_string())
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
            if let (Ok(old_range), Ok(new_range)) = (
                old_workbook.worksheet_formula(sheet),
                new_workbook.worksheet_formula(sheet),
            ) {
                let mut cell_diffs: Vec<CellDiff> = vec![];

                let (start_row, start_col, end_row, end_col) = diff_range(
                    old_range.start(),
                    new_range.start(),
                    old_range.end(),
                    new_range.end(),
                );

                for row in start_row..end_row {
                    for col in start_col..end_col {
                        let old_cell = match old_range.get_value((row, col)) {
                            Some(x) => &Data::String(x.to_string()),
                            None => &Data::Empty,
                        };
                        let new_cell = match new_range.get_value((row, col)) {
                            Some(x) => &Data::String(x.to_string()),
                            None => &Data::Empty,
                        };

                        if old_cell != new_cell {
                            let row = (row + 1) as usize;
                            let col = (col + 1) as usize;
                            cell_diffs.push(CellDiff {
                                row,
                                col,
                                addr: cell_pos_to_address(row, col),
                                kind: CellDiffKind::Formula,
                                old: if old_cell != &Data::Empty {
                                    Some(old_cell.to_string())
                                } else {
                                    None
                                },
                                new: if new_cell != &Data::Empty {
                                    Some(new_cell.to_string())
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

/// get range to compare
/// return: (start_row, start_col, end_row, end_col)
fn diff_range<'a>(
    old_start: Option<(u32, u32)>,
    new_start: Option<(u32, u32)>,
    old_end: Option<(u32, u32)>,
    new_end: Option<(u32, u32)>,
) -> (u32, u32, u32, u32) {
    let (old_start_row, old_start_col) = match old_start {
        Some((row, col)) => (row, col),
        None => (u32::MAX, u32::MAX),
    };
    let (new_start_row, new_start_col) = match new_start {
        Some((row, col)) => (row, col),
        None => (u32::MAX, u32::MAX),
    };
    let (old_end_row, old_end_col) = match old_end {
        Some((row, col)) => (row, col),
        None => (u32::MIN, u32::MIN),
    };
    let (new_end_row, new_end_col) = match new_end {
        Some((row, col)) => (row, col),
        None => (u32::MIN, u32::MIN),
    };
    let start_row = old_start_row.min(new_start_row);
    let start_col = old_start_col.min(new_start_col);
    let end_row = old_end_row.max(new_end_row);
    let end_col = old_end_col.max(new_end_col);

    (start_row, start_col, end_row + 1, end_col + 1)
}

/// convert (row, col) to cell address str
fn cell_pos_to_address(row: usize, col: usize) -> String {
    let col_letter = (col as u8 - 1) / 26;
    let col_index = (col as u8 - 1) % 26;

    let col_char = if col_letter == 0 {
        ((b'A' + col_index) as char).to_string()
    } else {
        let first_char = (b'A' + col_letter - 1) as char;
        let second_char = (b'A' + col_index) as char;
        format!("{}{}", first_char, second_char)
    };

    format!("{}{}", col_char, row)
}

/// unified diff line kind
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub enum UnifiedDiffKind {
    OldTitle,
    NewTitle,
    DiffPos,
    OldContent,
    NewContent,
}

/// unified diff lines
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct UnifiedDiff {
    pub lines: Vec<UnifiedDiffLine>,
}

impl fmt::Display for UnifiedDiff {
    /// to_string() for unified diff lines
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let str = self
            .lines
            .iter()
            .map(|x| x.to_string())
            .collect::<Vec<String>>();
        write!(f, "{}", str.join("\n"))
    }
}

/// unified diff line
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct UnifiedDiffLine {
    pub kind: UnifiedDiffKind,
    pub text: String,
}

impl fmt::Display for UnifiedDiffLine {
    /// to_string() for unified diff line
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self.kind {
            UnifiedDiffKind::OldTitle => write!(f, "--- {}", self.text),
            UnifiedDiffKind::NewTitle => write!(f, "+++ {}", self.text),
            UnifiedDiffKind::DiffPos => write!(f, "@@ {} @@", self.text),
            UnifiedDiffKind::OldContent => write!(f, "- {}", self.text),
            UnifiedDiffKind::NewContent => write!(f, "+ {}", self.text),
        }
    }
}
