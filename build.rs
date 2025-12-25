use std::env;

fn main() {
    // Solo compilar recursos en Windows
    if env::var("CARGO_CFG_TARGET_OS").unwrap() == "windows" {
        let mut res = winres::WindowsResource::new();
        res.set_icon("pdf_procu.ico");
        res.compile().expect("Error compilando recursos de Windows");
    }
}
