use std::{collections::HashMap, ops::RangeBounds};

use duplicate::duplicate_item;
use itertools::Itertools;

use crate::{re, utils::EnhanceVec};

pub fn parse_font_drawn(data: &str) -> HashMap<&str, &str> {
    let defs_regex = re!(r"<defs>.*?</defs>");
    let defs_tag = defs_regex.find(&data).unwrap().as_str();

    let path_regex = re!(r#"<path.*?id="(.*?)".*?d="(.*?)".*?/>"#);
    let mut font_dic = HashMap::new();
    path_regex.captures_iter(defs_tag).for_each(|c| {
        let font_name = c.get(1).unwrap().as_str();
        let d = c.get(2).unwrap().as_str();
        font_dic.insert(font_name, d);
    });

    font_dic
}

pub fn parse_details(data: &str) -> Vec<(String, &str, Vec<f64>)> {
    let g_regex = re!(r"<g.*?>(.*?)</g>");
    let g_tag = g_regex.find(&data).unwrap().as_str();
    let use_regex = re!(
        r##"<use.*?data-text="(.*?)".*?xlink:href="#(.*?)".*?transform="matrix\((.*?)\)".*?>"##
    );
    let use_list: Vec<_> = use_regex
        .captures_iter(g_tag)
        .map(|c| {
            let ch = c.get(1).unwrap().as_str().trim();
            let ch = if ch.starts_with("&#x") {
                let ch_part = &ch[3..ch.len() - 1];
                let codepoint = u32::from_str_radix(ch_part, 16).unwrap();
                char::from_u32(codepoint).unwrap().to_string().to_string()
            } else {
                ch.to_string()
            };

            let font_name = c.get(2).unwrap().as_str();

            let matrix: Vec<_> = c
                .get(3)
                .unwrap()
                .as_str()
                .split(",")
                .map(|each| each.parse::<f64>().unwrap())
                .collect();
            (ch, font_name, matrix)
        })
        .collect();

    use_list
}

#[derive(Debug)]
pub struct ChartSource {
    pub source: String,
    pub x_min: f64,
    pub x_max: f64,
    pub y: f64,
}

impl ChartSource {
    pub fn new(detail_list: &Vec<(String, &str, Vec<f64>)>, threshold: f64) -> Vec<ChartSource> {
        let mut res = vec![];
        let start_letter_regex = re!(r"^[a-zA-Z]");

        for (k, g) in &detail_list
            .iter()
            .filter(|&(_, _, mat)| mat[0] == 6.0)
            .group_by(|&(_, _, mat)| mat[5])
        {
            let mut p = g.peekable();

            let mut ch_res = String::new();
            let mut x_min = f64::MAX;
            let mut x_max = f64::MIN;
            loop {
                if let Some((ch_next, _, mat_next)) = p.next() {
                    ch_res.push_str(ch_next);
                    if mat_next[4] < x_min {
                        x_min = mat_next[4];
                    }
                    if mat_next[4] > x_max {
                        x_max = mat_next[4];
                    }

                    let mat_peek = match p.peek() {
                        Some((_, _, mat)) => mat,
                        None => mat_next,
                    };

                    if (mat_peek[4] - mat_next[4]).abs() > threshold {
                        if start_letter_regex.is_match(&ch_res) {
                            let cs = ChartSource {
                                source: ch_res.clone(),
                                x_min,
                                x_max,
                                y: k,
                            };
                            res.push(cs);
                        }
                        // 復位
                        ch_res.clear();
                        x_min = f64::MAX;
                        x_max = f64::MIN;
                    }
                } else {
                    if start_letter_regex.is_match(&ch_res) {
                        let cs = ChartSource {
                            source: ch_res.clone(),
                            x_min,
                            x_max,
                            y: k,
                        };
                        res.push(cs);
                    }
                    // 復位
                    ch_res.clear();
                    break;
                }
            }
        }

        res
    }
}

#[derive(Debug, Clone, Copy)]
pub struct ChartCode {
    pub codepoint: u32,
    pub hanzi: char,
    pub x_min: f64,
    pub x_max: f64,
    pub y: f64,
}

impl ChartCode {
    pub fn new(detail_list: &Vec<(String, &str, Vec<f64>)>, threshold: f64) -> Vec<ChartCode> {
        let mut res = vec![];

        for (k, g) in &detail_list
            .iter()
            .filter(|&(_, _, mat)| mat[0] == 9.9998)
            .group_by(|&(_, _, mat)| mat[5])
        {
            let mut p = g.peekable();

            let mut ch_res = String::new();
            let mut x_min = f64::MAX;
            let mut x_max = f64::MIN;
            loop {
                if let Some((ch_next, _, mat_next)) = p.next() {
                    ch_res.push_str(ch_next);
                    if mat_next[4] < x_min {
                        x_min = mat_next[4];
                    }
                    if mat_next[4] > x_max {
                        x_max = mat_next[4];
                    }

                    let mat_peek = match p.peek() {
                        Some((_, _, mat)) => mat,
                        None => mat_next,
                    };

                    if (mat_peek[4] - mat_next[4]).abs() > threshold {
                        if ch_res.len() >= 4 {
                            let codepoint = u32::from_str_radix(&ch_res, 16).unwrap();
                            let hanzi = char::from_u32(codepoint).unwrap();
                            let cs = ChartCode {
                                codepoint,
                                hanzi,
                                x_min,
                                x_max,
                                y: k,
                            };
                            res.push(cs);
                        }
                        // 復位
                        ch_res.clear();
                        x_min = f64::MAX;
                        x_max = f64::MIN;
                    }
                } else {
                    if ch_res.len() >= 4 {
                        let codepoint = u32::from_str_radix(&ch_res, 16).unwrap();
                        let hanzi = char::from_u32(codepoint).unwrap();
                        let cs = ChartCode {
                            codepoint,
                            hanzi,
                            x_min,
                            x_max,
                            y: k,
                        };
                        res.push(cs);
                    }
                    // 復位
                    ch_res.clear();
                    break;
                }
            }
        }

        res
    }

    pub fn hex(&self) -> String {
        format!("{:X}", self.codepoint)
    }
}

#[derive(Debug)]
pub struct ChartGraph<'a> {
    pub ch: char,
    pub drawn: &'a str,
    pub x: f64,
    pub y: f64,
}

impl<'b> ChartGraph<'b> {
    pub fn new<'a, R: RangeBounds<u32>>(
        detail_list: &Vec<(String, &str, Vec<f64>)>,
        font_dic: &'a HashMap<&str, &str>,
        codepoint_range: &R,
    ) -> Vec<ChartGraph<'a>> {
        let codepoint_begin = match codepoint_range.start_bound() {
            std::ops::Bound::Included(&num) => format!("\\u{{{:X}}}", num),
            std::ops::Bound::Excluded(&num) => format!("\\u{{{:X}}}", num + 1),
            std::ops::Bound::Unbounded => r"\u{0000}".to_string(),
        };
        let codepoint_end = match codepoint_range.end_bound() {
            std::ops::Bound::Included(&num) => format!("\\u{{{:X}}}", num),
            std::ops::Bound::Excluded(&num) => format!("\\u{{{:X}}}", num - 1),
            std::ops::Bound::Unbounded => r"\u{10FFFF}".to_string(),
        };
        let additional_range_str = format!(r"{}-{}", codepoint_begin, codepoint_end);
        let condition_str = r"[\u{3100}-\u{312F}\u{31A0}-\u{31BF}\u{4E00}-\u{9FFF}\u{3400}-\u{4DBF}\u{20000}-\u{2A6DF}\u{2A700}-\u{2B73F}\u{2B740}-\u{2B81F}\u{2B820}-\u{2CEAF}\u{2CEB0}-\u{2EBEF}\u{30000}-\u{3134F}\u{31350}-\u{323AF}\u{F900}-\u{FAFF}\u{2F800}-\u{2FA1F}\u{2F00}-\u{2FDF}\u{2E80}-\u{2EFF}\u{31C0}-\u{31EF}\u{2FF0}-\u{2FFF}\u{E000}-\u{F8FF}\u{F0000}-\u{FFFFD}\u{100000}-\u{10FFFD}".to_string() + &additional_range_str + "]";
        static RE: once_cell::sync::OnceCell<regex::Regex> = once_cell::sync::OnceCell::new();
        let character_regex = RE.get_or_init(|| {
            regex::RegexBuilder::new(&condition_str)
                .dot_matches_new_line(true)
                .build()
                .unwrap()
        });

        let res: Vec<_> = detail_list
            .iter()
            .filter(|&(ch, _, mat)| character_regex.is_match(ch) && mat[0] > 6.0)
            .map(|(ch, font_name, matrix)| ChartGraph {
                ch: ch.chars().next().unwrap(),
                drawn: font_dic[font_name],
                x: matrix[4],
                y: matrix[5],
            })
            .collect();

        res
    }
}

#[duplicate_item(
    [
        name [get_row_range_from_code_vec]
        method_1 [sort_x_y]
        method_2 [x_min]
    ]
    [
        name [get_col_range_from_code_vec]
        method_1 [sort_y_x]
        method_2 [y]
    ]
)]
pub fn name(code_vec: &mut Vec<ChartCode>) -> Vec<(f64, f64)> {
    code_vec.method_1(2.0);
    code_vec
        .iter()
        .map(|ele| ele.method_2)
        .dedup_by(|a, b| (a - b).abs() < 2.0)
        .chain([f64::MAX])
        .tuple_windows::<(_, _)>()
        .collect()
}

#[duplicate_item(
    [
        name [group_graph_source_iter_by_col]
        arg [row_ranges]
        method_1 [x_min]
    ]
    [
        name [group_graph_source_iter_by_row]
        arg [col_ranges]
        method_1 [y]
    ]
)]
pub fn name<'a, 'b>(
    graph_source_iter: impl Iterator<Item = (&'a ChartGraph<'a>, &'b ChartSource)>,
    arg: &Vec<(f64, f64)>,
) -> Vec<Vec<(&'a ChartGraph<'a>, &'b ChartSource)>> {
    let mut res_dic = HashMap::new();

    'outer: for (graph, source) in graph_source_iter {
        for (k, (start, end)) in arg.iter().enumerate() {
            if source.method_1 > *start && source.method_1 < *end {
                let tmp = res_dic.entry(k).or_insert(Vec::new());
                tmp.push((graph, source));
                continue 'outer;
            }
        }

        let tmp = res_dic.entry(arg.len()).or_insert(Vec::new());
        tmp.push((graph, source));
    }
    let res = res_dic
        .into_iter()
        .sorted_by_key(|&(a, _)| a)
        .filter(|(_, b)| !b.is_empty())
        .map(|(_, b)| b)
        .collect_vec();

    res
}

pub fn split_code_vec(code_vec: &mut Vec<ChartCode>) -> Vec<Vec<ChartCode>> {
    code_vec.sort_x_y(2.0);
    let mut res = vec![];
    for (_, v) in &code_vec.iter().group_by(|&ele| ele.x_min) {
        let tmp = v.map(|&each| each).collect_vec();
        res.push(tmp);
    }

    res
}
