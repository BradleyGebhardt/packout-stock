use rust_xlsxwriter::{Workbook, XlsxSerialize};
use serde::{Deserialize, Serialize};
use std::fs;
use std::fmt;

#[derive(Debug, Deserialize, Clone)]
struct Color {
    name: String,
    code: String,
}

#[derive(Debug, Deserialize, Clone)]
#[serde(tag = "type")]
enum ProductKind {
    Base {
        color: Color,
        size: i32,
    },
    Letter {
        color: Color,
        letter: char,
    },
    Icon { pack: i32 },
    Greenhouse { size: i32, price: u32 },
    Tumbler { name: String, price: u32, code: String },
    Dragon {
        name: String,
        size: Option<String>,
        code: String,
    },
}

impl std::fmt::Display for ProductKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.name())
    }
}

impl ProductKind {
    fn name(&self) -> String {
        let name = match self {
            ProductKind::Base { color, size, .. } => format!("Base {} {}", color.name, size),
            ProductKind::Letter { color, letter, .. } => format!("KC {} {}", color.name, letter),
            ProductKind::Icon { pack } => format!("Icon Pack {}", pack),
            ProductKind::Greenhouse { size, .. } => format!("Greenhouse {}", size),
            ProductKind::Tumbler { name, .. } => format!("Tumbler {}", name),
            ProductKind::Dragon { name, .. } => format!("Drg {}", name),
        };

        assert!(name.len() <= 15, "{} is over 15 chars", name);

        name
    }

    // CENTRALIZED SYSTEM: Every SKU mutation happens completely inside this single function block.
    fn sku(&self) -> String {
        let sku_base = "sayt";

        let sku = match self {
            ProductKind::Base { color, size, .. } => format!("{sku_base}-ab-{}-{size}", color.code),
            ProductKind::Letter { color, letter, .. } => format!("{sku_base}-al-{}-{letter}", color.code),
            ProductKind::Icon { pack } => format!("{sku_base}-pck-{}", pack),
            ProductKind::Greenhouse { size, .. } => format!("{sku_base}-gh-{size}"),
            ProductKind::Tumbler { code, .. } => format!("{sku_base}-{code}"),
            ProductKind::Dragon { code, .. } => format!("{sku_base}-{code}"),
        };

        assert!(sku.len() <= 15, "{}'s SKU is over 15 chars: {}, {}", self.name(), sku, sku.len());

        sku
    }

    fn price(&self) -> u32 {
        match self {
            ProductKind::Base { size, .. } => {
                let base = 75;
                base + (10 * (size - 3)) as u32
            }
            ProductKind::Letter { .. } => 6,
            ProductKind::Icon { .. } => 1,
            ProductKind::Greenhouse { price, .. } => *price,
            ProductKind::Tumbler { price, .. } => *price,
            ProductKind::Dragon { name, .. } => match name.as_str() {
                "Earth" => 150,
                "Small" => 150,
                "Medium" => 220,
                "Large" => 300,
                "Mini Fidget" => 60,
                _ => panic!("Dragon {} not covered", name),
            }
        }
    }
}

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

#[derive(Debug, Deserialize)]
struct ConfigFile {
    #[serde(default)]
    product: Vec<ProductKind>,
    #[serde(default)]
    matrix: Vec<MatrixRule>,
}

impl ConfigFile {
    fn flatten_to_catalog(self) -> Vec<ProductKind> {
        let mut full_catalog = self.product;

        for rule in self.matrix {
            match rule {
                MatrixRule::Letter { colors, range_start, range_end } => {
                    for color in &colors {
                        for c in (range_start as u8)..=(range_end as u8) {
                            let product = ProductKind::Letter {
                                color: color.clone(),
                                letter: c as char,
                            };
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
#[xlsx(header_format = Format::new().set_bold())]
struct FlattenedRow {
    #[xlsx(rename = "Name", column_width = 25.0)]
    name: String,
    #[xlsx(rename = "Price", column_width = 12.0)]
    price: u32,
    #[xlsx(rename = "SKU", column_width = 15.0)]
    sku: String,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let content = fs::read_to_string("products.toml")?;
    let config: ConfigFile = toml::from_str(&content)?;

    let final_catalog = config.flatten_to_catalog();
    println!("Dynamically generated {} products from configuration!", final_catalog.len());

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Stock")?;

    let rows: Vec<FlattenedRow> = final_catalog
        .iter()
        .map(|item| FlattenedRow {
            name: item.name(),
            price: item.price(),
            sku: item.sku(), // Call our single unified configuration pipeline
        })
        .collect();

    worksheet.set_serialize_headers::<FlattenedRow>(0, 0)?;
    worksheet.serialize(&rows)?;

    workbook.save("Stock.xlsx")?;

    Ok(())
}
