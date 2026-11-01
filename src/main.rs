use rust_xlsxwriter::{Format, Workbook, XlsxError, XlsxSerialize};
use serde::{Deserialize, Serialize};
use std::fs;

#[derive(Debug, Deserialize, Clone)]
struct Color {
    name: String,
    code: String,
}

// Represents individual parsed variants
#[derive(Debug, Deserialize, Clone)]
#[serde(tag = "type")]
enum ProductKind {
    Base { color: Color, size: i32 },
    Letter { color: Color, letter: char },
    Icon { pack: i32 },
    Greenhouse { size: i32 },
    Print { category: String, name: String, price: u32, sku: String },
    Tumbler { name: String, price: u32, sku: String },
}

impl ProductKind {
    fn name(&self) -> String {
        match self {
            ProductKind::Base { color, size } => format!("Base {} {}", color.name, size),
            ProductKind::Letter { color, letter } => format!("Keycap {} {}", color.code, letter),
            ProductKind::Icon { pack } => format!("Icon Pack {}", pack),
            ProductKind::Greenhouse { size } => format!("Greenhouse {}", size),
            ProductKind::Print { category, name, .. } => format!("{}, {}", category, name),
            ProductKind::Tumbler { name, .. } => format!("Tumbler {}", name),
        }
    }

    fn sku(&self) -> String {
        match self {
            ProductKind::Base { color, size } => format!("syab-{}-{}", color.code, size),
            ProductKind::Letter { color, letter } => format!("sykc-{}-{}", color.code, letter),
            ProductKind::Icon { pack } => format!("syipk-{}", pack),
            ProductKind::Greenhouse { size } => format!("sygh-{}", size),
            ProductKind::Print { sku, .. } => sku.clone(),
            ProductKind::Tumbler { sku, .. } => sku.clone(),
        }
    }

    fn price(&self) -> u32 {
        match self {
            ProductKind::Base { size, .. } => {
                let base = 75;
                base + (10 * (size - 3)) as u32
            }
            ProductKind::Letter { .. } => 6,
            ProductKind::Icon { .. } => 1,
            ProductKind::Greenhouse { size } => match size {
                4 => 100,
                6 => 120,
                _ => panic!("Greenhouse size {} price not handled", size),
            },
            ProductKind::Print { price, .. } => *price,
            ProductKind::Tumbler { price, .. } => *price,
        }
    }

    fn validate(&self) -> bool {
        self.sku().len() <= 9 && self.name().len() <= 15
    }
}

// Represents the rule-based expansion sets in the TOML file
#[derive(Debug, Deserialize)]
#[serde(tag = "type")]
enum MatrixRule {
    Letter {
        colors: Vec<Color>,
        range_start: char,
        range_end: char,
    },
    Base {
        colors: Vec<Color>,
        size_start: i32,
        size_end: i32,
    },
}

// The master root structure of your TOML file
#[derive(Debug, Deserialize)]
struct ConfigFile {
    #[serde(default)]
    product: Vec<ProductKind>, // The standalone entries
    #[serde(default)]
    matrix: Vec<MatrixRule>,   // The programmatic loops
}

impl ConfigFile {
    // Dynamically unpacks and flattens everything into a single clean catalog list
    fn flatten_to_catalog(self) -> Vec<ProductKind> {
        let mut full_catalog = self.product;

        for rule in self.matrix {
            match rule {
                MatrixRule::Letter { colors, range_start, range_end } => {
                    for color in &colors {
                        // Generate the char range safely
                        for c in (range_start as u8)..=(range_end as u8) {
                            let product = ProductKind::Letter {
                                color: color.clone(),
                                letter: c as char,
                            };
                            assert!(product.validate(), "Generated product failed validation: {:?}", product);
                            full_catalog.push(product);
                        }
                    }
                }
                MatrixRule::Base { colors, size_start, size_end } => {
                    for color in &colors {
                        for size in size_start..=size_end {
                            let product = ProductKind::Base {
                                color: color.clone(),
                                size,
                            };
                            assert!(product.validate(), "Generated product failed validation: {:?}", product);
                            full_catalog.push(product);
                        }
                    }
                }
            }
        }
        full_catalog
    }
}

#[derive(XlsxSerialize, Serialize)]
#[xlsx(header_format = Format::new().set_bold())] // Make header bold automatically!
struct FlattenedRow {
    #[xlsx(rename = "Name", column_width = 25.0)]
    name: String,
    
    #[xlsx(rename = "Price", column_width = 12.0)]
    price: u32,
    
    #[xlsx(rename = "SKU", column_width = 15.0)]
    sku: String,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Load and unpack the entire file layout smoothly
    let content = fs::read_to_string("products.toml")?;
    let config: ConfigFile = toml::from_str(&content)?;
    
    let final_catalog = config.flatten_to_catalog();
    println!("Dynamically generated {} products from your concise TOML configuration!", final_catalog.len());

    // You can now pipe `final_catalog` directly straight into your `FlattenedRow` mapping 
    // and call `worksheet.serialize(&rows)` just like before!

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Stock")?;

    // 2. Map polymorphic data into flat rows
    let rows: Vec<FlattenedRow> = final_catalog
        .iter()
        .map(|item| FlattenedRow {
            name: item.name(),
            price: item.price(),
            sku: item.sku(),
        })
        .collect();

    // 3. Dump the entire data structure straight into the worksheet instantly!
    worksheet.set_serialize_headers::<FlattenedRow>(0, 0)?; // Sets up headers automatically
    worksheet.serialize(&rows)?; // Dumps all data rows dynamically

    workbook.save("Stock.xlsx")?;

    Ok(())
}
