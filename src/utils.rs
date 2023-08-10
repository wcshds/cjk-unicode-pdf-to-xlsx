use duplicate::duplicate_item;

use crate::parse::{ChartCode, ChartSource, ChartGraph};

#[macro_export]
macro_rules! re {
    ($re:literal $(,)?) => {{
        static RE: once_cell::sync::OnceCell<regex::Regex> = once_cell::sync::OnceCell::new();
        RE.get_or_init(|| {
            regex::RegexBuilder::new($re)
                .dot_matches_new_line(true)
                .build()
                .unwrap()
        })
    }};
}

pub fn split_chart_source(raw_vec: &Vec<ChartSource>, threshold: f64) -> Vec<Vec<&ChartSource>> {
    let mut it = raw_vec.iter().peekable();
    let mut res = vec![];
    let mut tmp = vec![];
    loop {
        if let Some(cs) = it.next() {
            tmp.push(cs);
            let y = cs.y;
            let peek_y = match it.peek() {
                Some(ele) => ele.y,
                None => y,
            };

            if (peek_y - y).abs() > threshold {
                res.push(tmp);
                tmp = vec![];
            }
        } else {
            res.push(tmp);
            break;
        }
    }

    res
}

pub fn split_chart_graph<'a>(
    raw_vec: &'a Vec<ChartGraph<'a>>,
    threshold: f64,
) -> Vec<Vec<&'a ChartGraph<'a>>> {
    let mut it = raw_vec.iter().peekable();
    let mut res = vec![];
    let mut tmp = vec![];
    loop {
        if let Some(cs) = it.next() {
            tmp.push(cs);
            let y = cs.y;
            let peek_y = match it.peek() {
                Some(ele) => ele.y,
                None => y,
            };

            if (peek_y - y).abs() > threshold {
                res.push(tmp);
                tmp = vec![];
            }
        } else {
            res.push(tmp);
            break;
        }
    }

    res
}

pub trait EnhanceVec {
    fn sort_y_x(&mut self, threshold: f64);
    fn sort_x_y(&mut self, threshold: f64);
}

#[duplicate_item(
    [
        vec_type [Vec<ChartGraph<'_>>]
        method_x [x]
    ]
    [
        vec_type [Vec<ChartSource>]
        method_x [x_min]
    ]
    [
        vec_type [Vec<ChartCode>]
        method_x [x_min]
    ]
)]

impl EnhanceVec for vec_type {
    fn sort_y_x(&mut self, threshold: f64) {
        self.sort_by(|a, b| {
            if (a.y - b.y).abs() < threshold {
                a.method_x.partial_cmp(&b.method_x).unwrap()
            } else {
                a.y.partial_cmp(&b.y).unwrap()
            }
        });
    }

    fn sort_x_y(&mut self, threshold: f64) {
        self.sort_by(|a, b| {
            if (a.method_x - b.method_x).abs() < threshold {
                a.y.partial_cmp(&b.y).unwrap()
            } else {
                a.method_x.partial_cmp(&b.method_x).unwrap()
            }
        });
    }
}
