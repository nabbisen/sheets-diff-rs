use std::fmt;

#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

use super::diff::Diff;

/// unified diff
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct UnifiedDiff {
    pub content: Vec<UnifiedDiffContent>,
}

/// formatted unified diff
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct FormattedUnifiedDiff {
    pub content: Vec<UnifiedDiffContent>,
}

/// unified diff content
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct UnifiedDiffContent {
    pub old_title: String,
    pub new_title: String,
    pub lines: Vec<UnifiedDiffLine>,
}

/// unified diff content lines
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct UnifiedDiffLine {
    pub pos: Option<String>,
    pub old: Option<String>,
    pub new: Option<String>,
}

/// unified diff split into old / new parts
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct SplitUnifiedDiff {
    pub old: Vec<SplitUnifiedDiffContent>,
    pub new: Vec<SplitUnifiedDiffContent>,
}

/// unified diff content split into old / new parts
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct SplitUnifiedDiffContent {
    pub title: String,
    pub lines: Vec<SplitUnifiedDiffLine>,
}

/// unified diff content lines split into old / new parts
#[derive(Clone, Debug)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct SplitUnifiedDiffLine {
    pub pos: Option<String>,
    pub text: Option<String>,
}

impl UnifiedDiff {
    /// convert each string to one in unified format
    pub fn format(&self) -> FormattedUnifiedDiff {
        let content: Vec<UnifiedDiffContent> = self
            .content
            .iter()
            .map(|x| {
                let old_title = format!("--- {}", &x.old_title);
                let new_title = format!("+++ {}", &x.new_title);

                let lines: Vec<UnifiedDiffLine> = x
                    .lines
                    .iter()
                    .map(|x| {
                        let pos = if let Some(pos) = &x.pos {
                            Some(format!("@@ {} @@", pos))
                        } else {
                            None
                        };
                        let old = if let Some(old) = &x.old {
                            Some(format!("- {}", old))
                        } else {
                            None
                        };
                        let new = if let Some(new) = &x.new {
                            Some(format!("+ {}", new))
                        } else {
                            None
                        };
                        UnifiedDiffLine { pos, old, new }
                    })
                    .collect();

                UnifiedDiffContent {
                    old_title,
                    new_title,
                    lines,
                }
            })
            .collect();
        FormattedUnifiedDiff { content }
    }

    /// split into old / new parts
    pub fn split(&self) -> SplitUnifiedDiff {
        let old: Vec<SplitUnifiedDiffContent> = self
            .content
            .iter()
            .map(|x| {
                let title = x.old_title.clone();
                let lines: Vec<SplitUnifiedDiffLine> = x
                    .lines
                    .iter()
                    .map(|x| {
                        let pos = if let Some(pos) = &x.pos {
                            Some(pos.to_owned())
                        } else {
                            None
                        };
                        let text = if let Some(text) = &x.old {
                            Some(text.to_owned())
                        } else {
                            None
                        };
                        SplitUnifiedDiffLine { pos, text }
                    })
                    .collect();
                SplitUnifiedDiffContent { title, lines }
            })
            .collect();
        let new: Vec<SplitUnifiedDiffContent> = self
            .content
            .iter()
            .map(|x| {
                let title = x.new_title.clone();
                let lines: Vec<SplitUnifiedDiffLine> = x
                    .lines
                    .iter()
                    .map(|x| {
                        let pos = if let Some(pos) = &x.pos {
                            Some(pos.to_owned())
                        } else {
                            None
                        };
                        let text = if let Some(text) = &x.new {
                            Some(text.to_owned())
                        } else {
                            None
                        };
                        SplitUnifiedDiffLine { pos, text }
                    })
                    .collect();
                SplitUnifiedDiffContent { title, lines }
            })
            .collect();

        SplitUnifiedDiff { old, new }
    }
}

impl fmt::Display for FormattedUnifiedDiff {
    /// generate string in unified format
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        self.content.iter().for_each(|x| {
            let _ = writeln!(f, "{}", &x.old_title);
            let _ = writeln!(f, "{}", &x.new_title);
            x.lines.iter().for_each(|x| {
                if let Some(pos) = &x.pos {
                    let _ = writeln!(f, "{}", pos);
                }
                if let Some(old) = &x.old {
                    let _ = writeln!(f, "{}", old);
                }
                if let Some(new) = &x.new {
                    let _ = writeln!(f, "{}", new);
                }
            });
        });
        Ok(())
    }
}

/// get unified diff str split into old / new parts
pub fn unified_diff(diff: &Diff) -> UnifiedDiff {
    let mut ret: Vec<UnifiedDiffContent> = vec![];

    if !diff.sheet_diff.is_empty() {
        let old_title = format!("{} (sheet names)", diff.old_filepath);
        let new_title = format!("{} (sheet names)", diff.new_filepath);

        let lines: Vec<UnifiedDiffLine> = diff
            .sheet_diff
            .iter()
            .map(|x| {
                let old_sheet = if let Some(sheet) = x.old.as_ref() {
                    Some(sheet.to_owned())
                } else {
                    None
                };
                let new_sheet = if let Some(sheet) = x.new.as_ref() {
                    Some(sheet.to_owned())
                } else {
                    None
                };
                UnifiedDiffLine {
                    pos: None,
                    old: old_sheet,
                    new: new_sheet,
                }
            })
            .collect();

        ret.push(UnifiedDiffContent {
            old_title,
            new_title,
            lines,
        });
    }

    let cell_diffs_content: Vec<UnifiedDiffContent> = diff
        .cell_diffs
        .iter()
        .map(|x| {
            let cell_diffs_lines: Vec<UnifiedDiffLine> = x
                .cells
                .iter()
                .map(|x| {
                    let pos = Some(format!("{}({},{}) {}", x.addr, x.row, x.col, x.kind));

                    let old = if let Some(sheet) = x.old.as_ref() {
                        Some(sheet.to_owned())
                    } else {
                        None
                    };
                    let new = if let Some(sheet) = x.new.as_ref() {
                        Some(sheet.to_owned())
                    } else {
                        None
                    };

                    UnifiedDiffLine { pos, old, new }
                })
                .collect();

            UnifiedDiffContent {
                old_title: format!("{} [{}]", diff.old_filepath, x.sheet),
                new_title: format!("{} [{}]", diff.new_filepath, x.sheet),
                lines: cell_diffs_lines,
            }
        })
        .collect();

    ret.extend(cell_diffs_content);

    UnifiedDiff { content: ret }
}
