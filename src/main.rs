use dotext::{doc::MsDoc, doc::OpenOfficeDoc, *};
use std::{env, fs, io::Read};

fn main() {
    let input_path = &mut match env::args().nth(1) {
        Some(input_path) => input_path,
        None => panic!("Please provide a file path"),
    };

    let output_path = &mut match env::args().nth(2) {
        Some(output_path) => output_path,
        None => input_path.to_owned() + ".txt",
    };

    let mut text = String::new();

    if input_path.ends_with(".docx") {
        let mut docx = Docx::open(&input_path).unwrap();
        let _ = docx.read_to_string(&mut text);
    } else if input_path.ends_with(".odp") {
        let mut odp = Odp::open(&input_path).unwrap();
        let _ = odp.read_to_string(&mut text);
    } else if input_path.ends_with(".odt") {
        let mut odt = Odt::open(&input_path).unwrap();
        let _ = odt.read_to_string(&mut text);
    } else if input_path.ends_with(".pptx") {
        let mut pptx = Pptx::open(&input_path).unwrap();
        let _ = pptx.read_to_string(&mut text);
    } else if input_path.ends_with(".xlsx") {
        let mut xlsx = Xlsx::open(&input_path).unwrap();
        let _ = xlsx.read_to_string(&mut text);
    }

    fs::write(output_path, text).expect("Cannot write to file");
}
