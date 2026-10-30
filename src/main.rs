use rust_xlsxwriter::{Workbook, XlsxError, conditional_format::ConditionalFormat3ColorScale};

use std::fmt::{Display, Formatter, Result, format};

#[derive(Debug, Copy, Clone, Default)]
struct Color {
    name: &'static str,
    code: &'static str,
}

impl Display for Color {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(f, "Color:{name}|{code}", name = self.name, code = self.code)
    }
}

#[derive(Default)]
struct AlphaBase {
    size: i32,
    color: Color,
}

impl AlphaBase {
    fn name(&self) -> &str {
        "Base"
    }

    fn code(&self) -> &str {
        "ab"
    }

    fn sku(&self) -> String {
        format!("sy{}-{}-{}", self.code(), self.color.code, self.size)
    }

    fn price(&self) -> u32 {
        // 25% per item for the store
        let base: u32 = 75;

        if self.size == 3 {
            return base;
        }

        base + (10 * (self.size - 3)) as u32
    }
}

impl Display for AlphaBase {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "AlphaBase:{size}|{color}|{sku}",
            size = self.size,
            color = self.color,
            sku = self.sku()
        )
    }
}

struct AlphaLetter {
    letter: String,
    color: Color,
    base: Color,
}

struct AlphaIcon {
    icon: String,
    color: Color,
    base: Color,
}

#[derive(Default)]
struct AlphaPrebuilt {
    size: i32,
    color: Color,
}

impl AlphaPrebuilt {
    fn name(&self) -> &str {
        "Built"
    }

    fn code(&self) -> &str {
        "pb"
    }

    fn sku(&self) -> String {
        format!("sy{}-{}-{}", self.code(), self.color.code, self.size)
    }

    fn price(&self) -> u32 {
        // 25% per item for the store
        let base: u32 = 75;

        if self.size == 3 {
            return base;
        }

        base + (10 * (self.size - 3)) as u32
    }
}

impl Display for AlphaPrebuilt {
    fn fmt(&self, f: &mut Formatter<'_>) -> Result {
        write!(
            f,
            "AlphaPrebuilt:{size}|{color}|{sku}",
            size = self.size,
            color = self.color,
            sku = self.sku()
        )
    }
}


const LILAC_PURPLE: Color = Color {
    // name: "LPrp",
    name: "Purple",
    code: "lp",
};

const YELLOW: Color = Color {
    name: "Yellow",
    code: "y",
};

const BLACK: Color = Color {
    name: "Black",
    code: "bk",
};

const TURQOISE: Color = Color {
    name: "Turquoise",
    code: "tq",
};

const LIGHT_BLUE: Color = Color {
    name: "Ice Blue",
    code: "lb",
};

const PINK: Color = Color {
    name: "Pink",
    code: "pn",
};

const COLOR_ARRAY: [Color; 6] = [LILAC_PURPLE, YELLOW, BLACK, TURQOISE, LIGHT_BLUE, PINK];

// fn main() -> Result<(), XlsxError> {
fn main() -> Result {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let _ = worksheet.set_name("Stock");

    worksheet.write(0, 0, "name").expect("ERROR: Failed to write name header");
    worksheet.write(0, 1, "price").expect("ERROR: Failed to write price header");
    worksheet.write(0, 2, "sku").expect("ERROR: Failed to write sku header");

    let mut row = 1;

    for new_color in COLOR_ARRAY {
        for new_size in 3..8 {
            let base = AlphaBase {
                color: new_color.clone(),
                size: new_size,
                ..Default::default()
            };

            assert!(base.sku().len() <= 9, "{base}");
            assert!(base.name().len() <= 15, "{base}");

            worksheet.write(row, 0, format!("{} {} {}", base.name(), base.color.name, base.size)).expect("ERROR: Failed to write base name");
            worksheet.write(row, 1, format!("{}", base.price())).expect("ERROR: Failed to write base price");
            worksheet.write(row, 2, format!("{}", base.sku())).expect("ERROR: Failed to write base sku");

            row += 1;
        }
    };

    for new_color in COLOR_ARRAY {
        for new_size in 3..8 {
            let base = AlphaPrebuilt {
                color: new_color.clone(),
                size: new_size,
                ..Default::default()
            };

            assert!(base.sku().len() <= 9, "{base}");
            assert!(base.name().len() <= 15, "{base}");

            worksheet.write(row, 0, format!("{} {} {}", base.name(), base.color.name, base.size)).expect("ERROR: Failed to write base name");
            worksheet.write(row, 1, format!("{}", base.price())).expect("ERROR: Failed to write base price");
            worksheet.write(row, 2, format!("{}", base.sku())).expect("ERROR: Failed to write base sku");

            row += 1;
        }
    };

    workbook.save("Stock.xlsx").expect("ERROR: Failed to save Stock");

    Ok(())
}
