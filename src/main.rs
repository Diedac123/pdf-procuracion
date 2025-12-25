//! PDF Procuración - Rust Translation
//!
//! Este programa procesa archivos PDF y extrae información relevante como:
//! nombre, expediente, año, monto y número de cheque.
//! Luego guarda los datos en un archivo Excel y aplica formato.

use calamine::{open_workbook, Reader, Xlsx};
use lopdf::Document;
use regex::Regex;
use rfd::FileDialog;
use rust_xlsxwriter::{Formula, Table, TableColumn, TableStyle, Workbook};
use std::collections::HashSet;
use std::path::PathBuf;

/// Datos extraídos de una página del PDF
#[derive(Debug, Clone)]
struct DatosPagina {
    nombre: String,
    expediente: String,
    año: String,
    monto: String,
    cheque: String,
}

/// Extrae el texto entre comillas dobles que sigue a la palabra "autos".
fn extraer_texto_entre_comillas(texto: &str, p: usize) -> String {
    let patron = Regex::new(r#"autos\s+"(.*?)""#).unwrap();
    let eliminar_palabras: HashSet<&str> = ["ut-supra", "ut -supra"].iter().cloned().collect();

    let coincidencias: Vec<&str> = patron
        .captures_iter(texto)
        .filter_map(|c| c.get(1).map(|m| m.as_str()))
        .filter(|s| !eliminar_palabras.contains(*s))
        .collect();

    if coincidencias.is_empty() {
        (p + 1).to_string()
    } else {
        coincidencias[0].to_string()
    }
}

/// Extrae el número de expediente y el año del texto.
fn extraer_expediente_y_año(texto: &str, p: usize) -> (String, String) {
    // Normalizar texto
    let texto = Regex::new(r"(?i)expediente")
        .unwrap()
        .replace_all(texto, "EXP-")
        .to_string();
    let texto = Regex::new(r"(?i)Expte\.")
        .unwrap()
        .replace_all(&texto, "EXP-")
        .to_string();

    // Patrones sin lookbehind (incompatible con regex de Rust)
    let patrones_expediente = [
        r"[Ee][Xx][Pp]\-[^,]*,",
        r"[Ee][Xx][Pp]\.[^,]*,",
        r"[Ee][Xx][Pp] [^,]*,",
        r"\d{4,6}-\d{4}",
        r"\d{4,6}/\d{4}",
        r"[Ee][Jj][Ff]\-[^,]*,",
    ];

    let mut expediente: Option<String> = None;
    let mut patron_usado = "";

    for patron_str in &patrones_expediente {
        if let Ok(patron) = Regex::new(patron_str) {
            if let Some(m) = patron.find(&texto) {
                expediente = Some(m.as_str().to_uppercase().replace(' ', ""));
                patron_usado = patron_str;
                break;
            }
        }
    }

    let mut expediente = match expediente {
        Some(e) => e,
        None => return ((p + 1).to_string(), " ".to_string()),
    };

    // Para el caso de EXP #### y no EXP-####
    if patron_usado == r"[Ee][Xx][Pp] [^,]*," {
        expediente = expediente.replace("EXP", "");
    }

    // Acortar si es muy largo (usar chars para UTF-8 safety)
    let chars: Vec<char> = expediente.chars().collect();
    if chars.len() > 30 {
        expediente = chars[..25].iter().collect();
    }

    // Buscar el último dígito
    let chars: Vec<char> = expediente.chars().collect();
    let mut ultimo_digito = chars.len() as i32 - 1;
    while ultimo_digito >= 0 && !chars[ultimo_digito as usize].is_ascii_digit() {
        ultimo_digito -= 1;
    }
    if ultimo_digito >= 0 {
        expediente = chars[..=(ultimo_digito as usize)].iter().collect();
    }

    // Extraer año (usar chars para UTF-8 safety)
    let mut año = " ".to_string();
    let chars: Vec<char> = expediente.chars().collect();
    if chars.len() >= 4 {
        let año_str: String = chars[chars.len() - 4..].iter().collect();
        if let Ok(año_num) = año_str.parse::<i32>() {
            if (1990..=2025).contains(&año_num) {
                año = año_num.to_string();
                if chars.len() >= 5 {
                    expediente = chars[..chars.len() - 5].iter().collect();
                }
            }
        }
    }

    // Limpiar expediente
    expediente = expediente
        .replace("EXP. Nro.", "EXP-")
        .replace("Nº", "")
        .replace("N°", "")
        .replace("EXP.", "EXP-")
        .replace("NRO.", "");

    if expediente.ends_with('-') {
        expediente.pop();
    }

    if expediente.matches("EXP-").count() > 1 {
        expediente = expediente.replacen("EXP-", "", 1);
    } else if !expediente.contains("EXP-") && !expediente.contains("EJF-") {
        expediente = format!("EXP-{}", expediente);
    }

    (expediente, año)
}

/// Extrae el monto del texto.
fn extraer_monto(texto: &str, p: usize) -> String {
    let texto = texto.replace("( $", "($");

    // Buscar patrón ($...) sin lookbehind
    let patron = Regex::new(r"\(\$([^)]+)\)").unwrap();
    let coincidencia = match patron.captures(&texto) {
        Some(c) => c.get(1).map(|m| m.as_str()).unwrap_or(""),
        None => return (p + 1).to_string(),
    };

    let mut monto = coincidencia.replace('$', "").replace(' ', "");

    // Limpiar terminaciones
    if monto.ends_with(".-") {
        monto = monto[..monto.len() - 2].to_string();
    } else if monto.ends_with('.') {
        monto = monto[..monto.len() - 1].to_string();
    }

    if monto.len() < 3 {
        return monto;
    }

    let chars: Vec<char> = monto.chars().collect();
    let len = chars.len();
    let tercer_desde_final = chars.get(len.saturating_sub(3)).cloned().unwrap_or(' ');
    let num_puntos = monto.matches('.').count();
    let num_comas = monto.matches(',').count();

    if tercer_desde_final == '.' && num_puntos > 1 {
        // Caso: 1.234.567.89 -> 1234567.89
        let parte_decimal = &monto[monto.len() - 2..];
        let parte_entera = &monto[..monto.len() - 3];
        monto = format!("{}.{}", parte_entera.replace('.', ""), parte_decimal);
    } else if tercer_desde_final == '.' {
        // Caso: 1,234.56 -> 1234.56
        monto = monto.replace(',', "");
    } else if tercer_desde_final == ',' && num_comas > 1 {
        // Caso: 1,234,567,89 -> 1234567.89
        let parte_decimal = &monto[monto.len() - 2..];
        let parte_entera = &monto[..monto.len() - 3];
        monto = format!("{}.{}", parte_entera.replace(',', ""), parte_decimal);
    } else {
        // Caso: 1.234.567 o 1,234,567 -> entero
        monto = monto.replace('.', "").replace(',', ".");
    }

    monto
}

/// Extrae el número de cheque del texto.
fn extraer_numero_cheque(texto: &str, p: usize) -> String {
    let texto_limpio = texto.replace('.', "").replace('-', "").replace(' ', "");

    // Buscar ChequeNro o ChequeN°
    let patrones_cheque = [r"ChequeNro(\d+)", r"ChequeN°(\d+)"];
    for patron_str in &patrones_cheque {
        if let Ok(patron) = Regex::new(patron_str) {
            if let Some(caps) = patron.captures(&texto_limpio) {
                if let Some(m) = caps.get(1) {
                    let numero_str = m.as_str();
                    if numero_str.len() >= 8 {
                        let numero_str = &numero_str[..8];
                        if let Ok(numero) = numero_str.parse::<u64>() {
                            return format!("CH {}", numero);
                        }
                    }
                }
            }
        }
    }

    // Buscar ITBNº:
    if let Ok(patron) = Regex::new(r"ITBNº:(\d+)") {
        if let Some(caps) = patron.captures(&texto_limpio) {
            if let Some(m) = caps.get(1) {
                if let Ok(numero) = m.as_str().parse::<u64>() {
                    return format!("ITB {}", numero);
                }
            }
        }
    }

    // Buscar INTERNO:
    if let Ok(patron) = Regex::new(r"INTERNO:(\d+)") {
        if let Some(caps) = patron.captures(&texto_limpio) {
            if let Some(m) = caps.get(1) {
                let numero_str = m.as_str();
                if numero_str.len() > 4 {
                    let numero_str = &numero_str[..numero_str.len() - 4];
                    if let Ok(numero) = numero_str.parse::<u64>() {
                        if texto.contains("M.E.P.") {
                            return format!("MEP {}", numero);
                        }
                        return format!("ITB {}", numero);
                    }
                }
            }
        }
    }

    (p + 1).to_string()
}

/// Procesa un archivo PDF y extrae la información relevante de cada página.
fn procesar_pdf(ruta_archivo: &PathBuf) -> Result<Vec<DatosPagina>, Box<dyn std::error::Error>> {
    let doc = Document::load(ruta_archivo)?;
    let mut lista_datos = Vec::new();
    
    // Obtener todas las páginas del documento
    let pages = doc.get_pages();
    let num_pages = pages.len();
    
    println!("El PDF tiene {} páginas", num_pages);
    
    // Iterar sobre cada página
    for (p, (page_num, _page_id)) in pages.iter().enumerate() {
        // Extraer texto de esta página específica
        let texto_pagina = match doc.extract_text(&[*page_num]) {
            Ok(t) => t,
            Err(e) => {
                println!("Error extrayendo texto de página {}: {}", page_num, e);
                continue;
            }
        };
        
        // Limpiar el texto
        let texto: String = texto_pagina
            .replace('\n', " ")
            .chars()
            .filter(|c| c.is_ascii() || c.is_alphanumeric() || c.is_whitespace())
            .collect();
        
        // Saltar páginas con menos de 500 caracteres
        if texto.len() < 500 {
            println!("Página {} omitida: solo {} caracteres", page_num, texto.len());
            continue;
        }
        
        println!("Procesando página {} ({} caracteres)", page_num, texto.len());
        
        let nombre = extraer_texto_entre_comillas(&texto, p);
        let (expediente, año) = extraer_expediente_y_año(&texto, p);
        let monto = extraer_monto(&texto, p);
        let cheque = extraer_numero_cheque(&texto, p);
        
        lista_datos.push(DatosPagina {
            nombre,
            expediente,
            año,
            monto,
            cheque,
        });
    }
    
    Ok(lista_datos)
}

/// Guarda los datos en un archivo Excel y aplica formato.
fn guardar_y_formatear_excel(
    datos: &[DatosPagina],
    output_path: &PathBuf,
) -> Result<(), Box<dyn std::error::Error>> {
    // Leer datos existentes de la hoja REND si existe
    let datos_rend: Vec<Vec<String>> = if output_path.exists() {
        let mut workbook: Xlsx<_> = open_workbook(output_path)?;
        if let Ok(range) = workbook.worksheet_range("REND") {
            range
                .rows()
                .skip(1) // Skip header
                .map(|row| row.iter().map(|cell| cell.to_string()).collect())
                .collect()
        } else if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
            range
                .rows()
                .skip(1)
                .map(|row| row.iter().map(|cell| cell.to_string()).collect())
                .collect()
        } else {
            Vec::new()
        }
    } else {
        Vec::new()
    };

    let mut workbook = Workbook::new();

    // Crear hoja REND
    let worksheet_rend = workbook.add_worksheet();
    worksheet_rend.set_name("REND")?;

    // Encabezados REND
    let headers_rend = [
        "Numero de Cheque",
        "Monto",
        "AUTOS",
        "Expediente",
        "Año",
        "Observaciones",
        "Control",
        "Control cheque",
    ];

    for (col, header) in headers_rend.iter().enumerate() {
        worksheet_rend.write_string(0, col as u16, *header)?;
    }

    // Escribir datos REND existentes
    let mut max_row_rend = 0u32;
    for (row_idx, row_data) in datos_rend.iter().enumerate() {
        for (col_idx, cell) in row_data.iter().enumerate() {
            if col_idx < 8 {
                let row = (row_idx + 1) as u32;
                let col = col_idx as u16;
                
                // Columna B (índice 1) es Monto - escribir como número
                if col_idx == 1 {
                    // Convertir coma decimal a punto para parsear
                    let monto_normalizado = cell.replace(',', ".");
                    if let Ok(monto_num) = monto_normalizado.parse::<f64>() {
                        worksheet_rend.write_number(row, col, monto_num)?;
                    } else {
                        worksheet_rend.write_string(row, col, cell)?;
                    }
                } else {
                    worksheet_rend.write_string(row, col, cell)?;
                }
            }
        }
        max_row_rend = (row_idx + 1) as u32;
    }

    // Agregar fórmulas a REND (si hay datos)
    if max_row_rend > 0 {
        for row in 1..=max_row_rend {
            // D: Expediente
            let formula_d = format!(
                "=IFERROR(INDEX(PDF!$B:$B,MATCH(A{0},PDF!$E:$E,0)),INDEX(PDF!$B:$B,MATCH(B{0},PDF!$D:$D,0)))",
                row + 1
            );
            worksheet_rend.write_formula(row, 3, Formula::new(&formula_d))?;

            // E: Año
            let formula_e = format!(
                "=IFERROR(IF(INDEX(PDF!$C:$C,MATCH(A{0},PDF!$E:$E,0))>0,INDEX(PDF!$C:$C,MATCH(A{0},PDF!$E:$E,0)),\"\"),IF(INDEX(PDF!$C:$C,MATCH(B{0},PDF!$D:$D,0))>0,INDEX(PDF!$C:$C,MATCH(B{0},PDF!$D:$D,0)),\"\"))",
                row + 1
            );
            worksheet_rend.write_formula(row, 4, Formula::new(&formula_e))?;

            // G: Control
            let formula_g = format!("=COUNTIF(PDF!$D:$D,B{})", row + 1);
            worksheet_rend.write_formula(row, 6, Formula::new(&formula_g))?;

            // H: Control cheque
            let formula_h = format!("=COUNTIF(PDF!$E:$E,A{})", row + 1);
            worksheet_rend.write_formula(row, 7, Formula::new(&formula_h))?;
        }

        // Crear tabla REND
        let table_rend = Table::new()
            .set_style(TableStyle::Light1)
            .set_columns(&[
                TableColumn::new().set_header("Numero de Cheque"),
                TableColumn::new().set_header("Monto"),
                TableColumn::new().set_header("AUTOS"),
                TableColumn::new().set_header("Expediente"),
                TableColumn::new().set_header("Año"),
                TableColumn::new().set_header("Observaciones"),
                TableColumn::new().set_header("Control"),
                TableColumn::new().set_header("Control cheque"),
            ]);
        worksheet_rend.add_table(0, 0, max_row_rend, 7, &table_rend)?;
    }

    // Crear hoja PDF
    let worksheet_pdf = workbook.add_worksheet();
    worksheet_pdf.set_name("PDF")?;

    // Encabezados PDF
    let headers_pdf = [
        "Nombre",
        "Expediente",
        "año",
        "Monto",
        "Cheque",
        "Control",
        "Control cheque",
    ];

    for (col, header) in headers_pdf.iter().enumerate() {
        worksheet_pdf.write_string(0, col as u16, *header)?;
    }

    // Escribir datos extraídos del PDF
    for (row_idx, dato) in datos.iter().enumerate() {
        let row = (row_idx + 1) as u32;

        worksheet_pdf.write_string(row, 0, &dato.nombre)?;
        worksheet_pdf.write_string(row, 1, &dato.expediente)?;

        // Escribir año como número si es posible
        if let Ok(año_num) = dato.año.trim().parse::<f64>() {
            worksheet_pdf.write_number(row, 2, año_num)?;
        } else {
            worksheet_pdf.write_string(row, 2, &dato.año)?;
        }

        // Escribir monto como número si es posible
        if let Ok(monto_num) = dato.monto.parse::<f64>() {
            worksheet_pdf.write_number(row, 3, monto_num)?;
        } else {
            worksheet_pdf.write_string(row, 3, &dato.monto)?;
        }

        worksheet_pdf.write_string(row, 4, &dato.cheque)?;

        // Fórmulas de control
        let formula_f = format!("=COUNTIF(REND!$B:$B,D{})", row + 1);
        let formula_g = format!("=COUNTIF(REND!$A:$A,E{})", row + 1);
        worksheet_pdf.write_formula(row, 5, Formula::new(&formula_f))?;
        worksheet_pdf.write_formula(row, 6, Formula::new(&formula_g))?;
    }

    // Crear tabla PDF
    if !datos.is_empty() {
        let max_row_pdf = datos.len() as u32;
        let table_pdf = Table::new()
            .set_style(TableStyle::Light1)
            .set_columns(&[
                TableColumn::new().set_header("Nombre"),
                TableColumn::new().set_header("Expediente"),
                TableColumn::new().set_header("año"),
                TableColumn::new().set_header("Monto"),
                TableColumn::new().set_header("Cheque"),
                TableColumn::new().set_header("Control"),
                TableColumn::new().set_header("Control cheque"),
            ]);
        worksheet_pdf.add_table(0, 0, max_row_pdf, 6, &table_pdf)?;
    }

    workbook.save(output_path)?;
    Ok(())
}

fn main() {
    println!("PDF Procuración - Procesador de PDFs");
    println!("=====================================\n");

    // Seleccionar archivo PDF
    println!("Seleccione el archivo PDF a procesar...");
    let pdf_file = FileDialog::new()
        .add_filter("Archivos PDF", &["pdf"])
        .set_title("Seleccionar archivo PDF")
        .pick_file();

    let pdf_path = match pdf_file {
        Some(path) => path,
        None => {
            println!("No se seleccionó ningún archivo PDF.");
            return;
        }
    };

    println!("Procesando: {:?}", pdf_path);

    // Procesar PDF
    let datos = match procesar_pdf(&pdf_path) {
        Ok(d) => d,
        Err(e) => {
            println!("Error al procesar el PDF: {}", e);
            return;
        }
    };

    println!("Se extrajeron {} registros del PDF.", datos.len());

    if datos.is_empty() {
        println!("No se encontraron datos en el PDF.");
        return;
    }

    // Seleccionar archivo Excel de salida
    println!("\nSeleccione el archivo Excel de salida...");
    let excel_file = FileDialog::new()
        .add_filter("Archivos Excel", &["xlsx"])
        .set_title("Seleccionar archivo Excel")
        .pick_file();

    let excel_path = match excel_file {
        Some(path) => path,
        None => {
            println!("No se seleccionó ningún archivo Excel.");
            return;
        }
    };

    // Guardar y formatear Excel
    match guardar_y_formatear_excel(&datos, &excel_path) {
        Ok(_) => println!(
            "\n✓ Archivo Excel guardado y formateado correctamente: {:?}",
            excel_path
        ),
        Err(e) => println!("Error al guardar el archivo Excel: {}", e),
    }
}
