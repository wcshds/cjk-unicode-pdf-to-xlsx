use cjk_unicode_pdf_to_xlsx::run;

fn main() {
    run("./cjk-unicode-pdf/U4E00.pdf", "./result/basic-1.xlsx", 1..=300, 0x4e00..=0x9fff, 100).unwrap();
    run("./cjk-unicode-pdf/U4E00.pdf", "./result/basic-2.xlsx", 301..=532, 0x4e00..=0x9fff, 100).unwrap();

    // cannot open
    run("./cjk-unicode-pdf/U4E00.pdf", "./result/basic-full.xlsx", 1..=532, 0x4e00..=0x9fff, 100).unwrap();
}
