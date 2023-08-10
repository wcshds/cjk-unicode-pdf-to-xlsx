use itertools::Itertools;
use pyo3::{types::IntoPyDict, PyResult, Python};
use std::ops::RangeBounds;

use image_gen::svg_drawn_to_image;
use parse::*;
use utils::*;
use xlsx::Xlsx;

pub mod image_gen;
pub mod parse;
pub mod utils;
pub mod xlsx;

pub fn run<R1: RangeBounds<usize> + Iterator<Item = usize>, R2: RangeBounds<u32>>(
    filename: &str,
    output: &str,
    page_range: R1,
    codepoint_range: R2,
    limit: u32,
) -> PyResult<()> {
    let mut xlsx = Xlsx::new();

    Python::with_gil(|py| {
        let fitz = py.import("fitz")?;

        let doc = fitz.call_method1("open", (filename,))?;
        let fitz_matrix_identity = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0);

        let mut page_count = 1;
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
                panic!("解析文件 {} 的第 {} 頁時發生了錯誤...", filename, page_idx);
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

        xlsx.save(output);
        Ok(())
    })
}
