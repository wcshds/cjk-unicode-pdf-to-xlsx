use std::{io::Cursor, path::Path};

use image::{ImageFormat, Luma};
use once_cell::sync::Lazy;
use rust_xlsxwriter::{Format, Image, Workbook};

pub struct Xlsx {
    workbook: Workbook,
    current_row: u32,
    current_sheet: usize,
}

impl Xlsx {
    pub fn new() -> Self {
        let mut workbook = Workbook::new();
        workbook.add_worksheet();
        Self {
            workbook,
            current_row: 0,
            current_sheet: 0,
        }
    }

    pub fn add_format(&mut self, current_row: u32, col_count: usize, col_max: usize) {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.current_sheet)
            .unwrap();

        for col in (col_count + 1)..=col_max {
            worksheet
                .write_with_format(current_row, col as u16, "", &MIDDLE_TOP_FORMAT)
                .unwrap();
            worksheet
                .write_with_format(current_row + 1, col as u16, "", &MIDDLE_MIDDLE_FORMAT)
                .unwrap();
            worksheet
                .write_with_format(current_row + 2, col as u16, "", &MIDDLE_BOTTOM_FORMAT)
                .unwrap();
        }

        if col_count != col_max {
            worksheet
                .write_with_format(current_row, col_max as u16, "", &LAST_TOP_FORMAT)
                .unwrap();
        }

        worksheet
            .write_with_format(current_row + 1, col_max as u16, "", &LAST_MIDDLE_FORMAT)
            .unwrap();
        worksheet
            .write_with_format(current_row + 2, col_max as u16, "", &LAST_BOTTOM_FORMAT)
            .unwrap();
    }

    pub fn add_row<S: AsRef<str>>(
        &mut self,
        codepoint_hex: &str,
        images_with_sources: &Vec<(S, image::ImageBuffer<Luma<u8>, Vec<u8>>)>,
        col_max: usize,
    ) {
        let worksheet = self
            .workbook
            .worksheet_from_index(self.current_sheet)
            .unwrap();

        worksheet
            .set_row_height_pixels(self.current_row + 1, 85)
            .unwrap();

        worksheet
            .merge_range(
                self.current_row,
                0,
                self.current_row + 2,
                0,
                codepoint_hex,
                &FIRST_FORMAT,
            )
            .unwrap();

        for (col, (source, image_origin)) in (1..).zip(images_with_sources) {
            if col == col_max as u16 {
                worksheet
                    .write_with_format(self.current_row, col, source.as_ref(), &LAST_TOP_FORMAT)
                    .unwrap();
            } else {
                worksheet
                    .write_with_format(self.current_row, col, source.as_ref(), &MIDDLE_TOP_FORMAT)
                    .unwrap();
            }
            let mut buf = Cursor::new(vec![]);
            image_origin.write_to(&mut buf, ImageFormat::Png).unwrap();
            let mut image = Image::new_from_buffer(buf.get_ref()).unwrap();
            image.set_scale_width(0.65).set_scale_height(0.65);

            worksheet.set_column_width_pixels(col, 85).unwrap();
            worksheet
                .insert_image_with_offset(self.current_row + 1, col, &image, 1, 1)
                .unwrap();

            // 設置格式
            worksheet
                .write_with_format(self.current_row + 1, col, "", &MIDDLE_MIDDLE_FORMAT)
                .unwrap();
            worksheet
                .write_with_format(self.current_row + 2, col, "", &MIDDLE_BOTTOM_FORMAT)
                .unwrap();
        }

        self.add_format(self.current_row, images_with_sources.len(), col_max);

        self.current_row += 3;
    }

    pub fn next_sheet(&mut self) {
        self.workbook.add_worksheet();
        self.current_sheet += 1;
        self.current_row = 0;
    }

    pub fn save<P: AsRef<Path>>(&mut self, path: P) {
        self.workbook.save(path).unwrap();
    }

    pub fn save_to_buffer(&mut self) -> Vec<u8> {
        self.workbook.save_to_buffer().unwrap()
    }
}

static FIRST_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
static MIDDLE_TOP_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
static MIDDLE_MIDDLE_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
static MIDDLE_BOTTOM_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thin)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
static LAST_TOP_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thick)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
static LAST_MIDDLE_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thick)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
static LAST_BOTTOM_FORMAT: Lazy<Format> = Lazy::new(|| {
    Format::new()
        .set_border_top(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_left(rust_xlsxwriter::FormatBorder::Thin)
        .set_border_bottom(rust_xlsxwriter::FormatBorder::Thick)
        .set_border_right(rust_xlsxwriter::FormatBorder::Thick)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
});
