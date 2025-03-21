#[cfg(test)]
mod tests {
    use sheets_diff::core::unified_format::unified_diff;

    #[test]
    fn it_works() {
        const OLD_FILEPATH: &str = "tests/fixtures/file1.xlsx";
        const NEW_FILEPATH: &str = "tests/fixtures/file2.xlsx";

        const EXPECT: &str = r#"--- tests/fixtures/file1.xlsx (sheet names)
+++ tests/fixtures/file2.xlsx (sheet names)
- Sheet1_2
+ Sheetzz
--- tests/fixtures/file1.xlsx [Sheet1]
+++ tests/fixtures/file2.xlsx [Sheet1]
@@ A1(1,1) value @@
- 1
@@ B2(2,2) value @@
- 2
+ 今日は世界
@@ B4(4,2) value @@
+ a
@@ C6(6,3) value @@
+ hej
@@ D10(10,4) value @@
- 2
+ 8
@@ D10(10,4) formula @@
- 1+1
+ 2*4
@@ D11(11,4) formula @@
+ 
@@ D12(12,4) value @@
+ a123
@@ D12(12,4) formula @@
+ "a"&123
@@ W55(55,23) value @@
+ っｓ
"#;

        let diff = sheets_diff::core::diff::Diff::new(OLD_FILEPATH, NEW_FILEPATH);
        let target = unified_diff(&diff);
        assert_eq!(format!("{}", target), EXPECT);
    }
}
