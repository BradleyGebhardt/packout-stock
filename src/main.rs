use calamine::{self, Reader};
use rust_xlsxwriter::{Workbook, XlsxSerialize};
use serde;
use std::collections::HashMap;
use strfmt::strfmt;
use barcoders::sym::code39::Code39;
use barcoders::generators::image::Image;
use std::fs::File;
use std::io::BufWriter;
use std::path::Path;

const DEFAULT_SKU_FORMAT: &'static str = "sy-{code}";

#[derive(Debug, serde::Deserialize)]
struct RawStock {
    name: String,
    price: u32,
    code: String,
    #[serde(rename = "sku-format", default)]
    sku_format: Option<String>,
    extra_one: Option<String>,
    extra_two: Option<String>,
    extra_three: Option<String>,
}

#[derive(XlsxSerialize, serde::Serialize)]
struct StockOutput {
    #[xlsx(rename = "Name", column_width = 25.0)]
    name: String,
    #[xlsx(rename = "Price", column_width = 15.0)]
    price: u32,
    #[xlsx(rename = "Sku", column_width = 25.0)]
    sku: String,
}

impl std::fmt::Display for RawStock {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let sku_format = self
            .sku_format
            .clone()
            .unwrap_or(DEFAULT_SKU_FORMAT.to_string());
        let sku = self.sku();
        write!(
            f,
            "{}, {}, {}, {}, {}",
            self.name, self.code, self.price, sku_format, sku
        )
    }
}

impl RawStock {
    fn to_format_map(&self) -> HashMap<String, String> {
        let mut map = HashMap::new();

        map.insert("name".to_string(), self.name.clone());
        map.insert("code".to_string(), self.code.clone());
        map.insert("price".to_string(), self.price.to_string());
        map.insert(
            "extra_one".to_string(),
            self.extra_one.clone().unwrap_or("".to_string()),
        );
        map.insert(
            "extra_two".to_string(),
            self.extra_two.clone().unwrap_or("".to_string()),
        );
        map.insert(
            "extra_three".to_string(),
            self.extra_three.clone().unwrap_or("".to_string()),
        );

        map
    }

    fn sku_format(&self) -> String {
        self.sku_format
            .clone()
            .unwrap_or(DEFAULT_SKU_FORMAT.to_string())
    }

    fn sku(&self) -> String {
        let vars = self.to_format_map();
        let sku_format = self.sku_format();

        match strfmt(&sku_format, &vars) {
            Ok(rendered) => rendered.to_ascii_uppercase(),
            Err(e) => panic!("ERROR: failed to format sku: {e}"),
        }
    }

    fn validate(&self) {
        let sku = self.sku();
        let sku_length = sku.len();
        assert!(
            sku_length <= 9,
            "ERROR: {}'s SKU is over 9 chars; {}, {}, {} chars",
            self.name,
            self.sku_format(),
            sku,
            sku_length,
        );

        let name_length = self.name.len();

        assert!(
            name_length <= 15,
            "ERROR: {} is over 15 chars, {}",
            self.name,
            name_length
        );
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut raw_workbook: calamine::Xlsx<_> = calamine::open_workbook("Stock.raw.xlsx")?;
    let raw_stock_worksheet = raw_workbook.worksheet_range("Stock")?;
    let raw_stock = calamine::RangeDeserializerBuilder::with_deserialize_headers::<RawStock>()
        .from_range::<_, RawStock>(&raw_stock_worksheet)?;

    let mut stock_list: Vec<StockOutput> = Vec::new();

    for row in raw_stock {
        let row = row?;
        row.validate();

        let barcode = Code39::new(row.sku())?;

        let encoded_barcode = barcode.encode();

        let barcode_png = Image::png(80);
        let png_bytes = barcode_png.generate(encoded_barcode)?;

        stock_list.push(StockOutput {
            name: row.name.clone(),
            price: row.price,
            sku: row.sku(),
        });
    }

    let mut output_workbook = Workbook::new();
    let output_worksheet = output_workbook.add_worksheet();
    output_worksheet.set_name("Stock")?;

    output_worksheet.set_serialize_headers::<StockOutput>(0, 0)?; // Sets up headers automatically
    output_worksheet.serialize(&stock_list)?; // Dumps all data rows dynamically

    output_workbook.save("Stock.xlsx")?;

    Ok(())
}
