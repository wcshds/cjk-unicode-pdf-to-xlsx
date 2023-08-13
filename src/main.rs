use cjk_unicode_pdf_to_xlsx::run;

fn main() {
    let infos = vec![
        ("U4E00.pdf", "basic.xlsx", 0x4e00..=0x9fff),
        // ("U3400.pdf", "ext-a.xlsx", 0x3400..=0x4DBF),
        ("U20000.pdf", "ext-b.xlsx", 0x20000..=0x2A6DF),
        // ("U2A700.pdf", "ext-c.xlsx", 0x2A700..=0x2B739),
        // ("U2B740.pdf", "ext-d.xlsx", 0x2B740..=0x2B81D),
        // ("U2B820.pdf", "ext-e.xlsx", 0x2B820..=0x2CEA1),
        // ("U2CEB0.pdf", "ext-f.xlsx", 0x2CEB0..=0x2EBE0),
        // ("U30000.pdf", "ext-g.xlsx", 0x30000..=0x3134A),
        // ("U31350.pdf", "ext-h.xlsx", 0x31350..=0x323AF),
    ];

    for (input_name, output_name, codepoint_range) in infos {
        let input_path = format!("./cjk-unicode-pdf/{}", input_name);
        let output_path = format!("./result/{}", output_name);
        println!("正在處理文件 {} 中:", input_path);
        run(&input_path, &output_path, 1.., codepoint_range, 100).unwrap();
        println!();
    }
}
