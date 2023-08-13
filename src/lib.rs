use itertools::Itertools;
use pyo3::{
    types::{IntoPyDict, PyDict},
    PyResult, Python,
};
use std::{io::Cursor, ops::RangeBounds, path::Path};

use image_gen::svg_drawn_to_image;
use parse::*;
use utils::*;
use xlsx::Xlsx;

pub mod image_gen;
pub mod parse;
pub mod utils;
pub mod xlsx;

pub fn run<R1: RangeBounds<usize> + Iterator<Item = usize>, R2: RangeBounds<u32>>(
    input: &str,
    output: &str,
    page_range: R1,
    codepoint_range: R2,
    limit: u32,
) -> PyResult<()> {
    let mut xlsx = Xlsx::new();

    Python::with_gil(|py| {
        let fitz = py.import("fitz")?;

        let doc = fitz.call_method1("open", (input,))?;
        let fitz_matrix_identity = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0);

        let pdf_len: usize = doc.call_method0("__len__")?.extract()?; // 首頁爲說明頁

        let mut page_count = 1;
        let page_range = page_range_normalize(page_range, 1, pdf_len - 1);
        let mut page_range_peek = page_range.peekable();
        while let Some(page_idx) = page_range_peek.next() {
            let page = doc.call_method1("__getitem__", (page_idx,))?;
            let page_svg: String = page
                .call_method(
                    "get_svg_image",
                    (),
                    Some(vec![("matrix", &fitz_matrix_identity)].into_py_dict(py)),
                )?
                .extract()?;

            // 初步解析
            let font_dic = parse_font_drawn(&page_svg);
            let detail_list = parse_details(&page_svg); // character, font-name, matrix

            // 結構化
            let mut source_vec = ChartSource::new(&detail_list, 7.0);
            let mut graph_vec = ChartGraph::new(&detail_list, &font_dic, &codepoint_range);
            let mut code_vec = ChartCode::new(&detail_list, 10.0); // 整頁的code_vec

            // 排序
            source_vec.sort_y_x(5.0);
            graph_vec.sort_y_x(5.0);

            // 按行分割
            let source_rows = split_chart_source(&source_vec, 10.0);
            let graph_rows = split_chart_graph(&graph_vec, 10.0);

            if !(source_vec.len() == graph_vec.len() && source_rows.len() == graph_rows.len()) {
                panic!("解析文件 {} 的第 {} 頁時發生了錯誤...", input, page_idx);
            }

            // 單列的code_vec
            let mut code_vec_split = split_code_vec(&mut code_vec);
            // 依照編碼的位置將 (字形, 字源) 按列分組
            let row_ranges = get_row_range_from_code_vec(&mut code_vec);
            let cols =
                group_graph_source_iter_by_col(graph_vec.iter().zip(&source_vec), &row_ranges);

            let col_max = 7;

            for (code_vec_each, col_each) in code_vec_split.iter_mut().zip(cols.into_iter()) {
                // 將每列的 (字形, 字源) 再按行分組
                let col_ranges = get_col_range_from_code_vec(code_vec_each);
                let elements = group_graph_source_iter_by_row(col_each.into_iter(), &col_ranges);

                for (code, graph_source) in code_vec_each.iter().zip(elements.into_iter()) {
                    // println!("{}", code.hex());

                    let images_with_source = graph_source
                        .into_iter()
                        .map(|(graph, source)| {
                            let img = svg_drawn_to_image(graph.drawn);

                            (source.source.clone(), img)
                        })
                        .collect_vec();

                    xlsx.add_row(&code.hex(), &images_with_source, col_max);
                }
            }

            println!("第 {:03} 頁已處理", page_idx);

            if page_count % limit == 0 && page_range_peek.peek().is_some() {
                xlsx.next_sheet();
            }
            page_count += 1;
        }

        // xlsx.save(output);
        // Ok(())

        let zip_buf = xlsx.save_to_buffer();
        match rezip(&py, &zip_buf, output) {
            Ok(()) => Ok(()),
            Err(err) => Err(err),
        }
    })
}

fn rezip<P: AsRef<Path>>(py: &pyo3::Python, zip_buf: &Vec<u8>, path: P) -> PyResult<()> {
    let zipfile_class = py.import("zipfile").unwrap().getattr("ZipFile")?;
    let kwargs = PyDict::new(*py);
    kwargs.set_item("mode", "w")?;
    kwargs.set_item("compression", 8)?;
    kwargs.set_item("compresslevel", 6)?;
    let zipfile = zipfile_class
        .call((path.as_ref().to_str().unwrap(),), Some(kwargs))
        .unwrap();

    let reader = Cursor::new(zip_buf);
    let mut zip_archive = zip::ZipArchive::new(reader).unwrap();

    let mut buf_writer = vec![];

    for i in 0..zip_archive.len() {
        let mut file = zip_archive.by_index(i).unwrap();
        let filepath = match file.enclosed_name() {
            Some(path) => path.to_owned(),
            None => continue,
        };

        std::io::copy(&mut file, &mut buf_writer).unwrap();

        zipfile
            .call_method1("writestr", (filepath.to_str().unwrap(), &buf_writer[..]))
            .unwrap();

        buf_writer.clear();
    }

    zipfile.call_method0("close").unwrap();

    Ok(())
}

fn page_range_normalize<R: RangeBounds<usize> + Iterator<Item = usize>>(
    page_range: R,
    min: usize,
    max: usize,
) -> std::ops::Range<usize> {
    let start = match page_range.start_bound() {
        std::ops::Bound::Included(&num) => num,
        std::ops::Bound::Excluded(_) => min,
        std::ops::Bound::Unbounded => min,
    };
    let end = match page_range.end_bound() {
        std::ops::Bound::Included(&num) => num + 1,
        std::ops::Bound::Excluded(&num) => num,
        std::ops::Bound::Unbounded => max + 1,
    };

    start..end
}
