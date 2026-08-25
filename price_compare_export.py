import io

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_col_to_name


def create_excel(
    full_result,
    supplier_price_columns,
    supplier_configs,
    duplicate_info,
    our_file_name,
    our_config,
):
    export_result = full_result.drop(columns=["_relevant_for_display"]).copy()

    if our_config["type"] == "PowerBI Pricelist":
        browser_count = int(full_result["_relevant_for_display"].sum())
        browser_metric = "Relevant products shown in browser"
    else:
        browser_count = len(full_result)
        browser_metric = "Products shown in browser"

    summary_rows = [
        ("Our pricelist type", our_config["type"]),
        ("OurPrices file", our_file_name),
        ("Suppliers uploaded", len(supplier_configs)),
        ("All current-pricelist products exported", len(full_result)),
        (browser_metric, browser_count),
        (
            "Products matched with at least one supplier",
            int(full_result["Matched Suppliers"].gt(0).sum()),
        ),
        (
            "Products where at least one supplier is cheaper",
            int((full_result["Status"] == "CHEAPER").sum()),
        ),
        (
            "Products where cheapest supplier is more expensive",
            int((full_result["Status"] == "MORE EXPENSIVE").sum()),
        ),
        (
            "Products with same cheapest price",
            int((full_result["Status"] == "SAME PRICE").sum()),
        ),
        (
            "Products not found at any supplier",
            int((full_result["Status"] == "NOT FOUND").sum()),
        ),
    ]

    if our_config["type"] == "Free format pricelist":
        summary_rows.extend(
            [
                ("", ""),
                ("Free format - EAN column", our_config.get("ean_column") or "Not selected"),
                ("Free format - SKU column", our_config.get("sku_column") or "Not selected"),
                ("Free format - Price column", our_config.get("price_column")),
                (
                    "Free format - Additional columns",
                    ", ".join(map(str, our_config.get("extra_columns", []))) or "None",
                ),
            ]
        )

    for config in supplier_configs:
        supplier_name = config["supplier_name"]
        price_col = f"{supplier_name} Price"
        supplier_rows = [
            ("", ""),
            (f"{supplier_name} - source file", config["file_name"]),
            (f"{supplier_name} - sheet", config["sheet_name"]),
            (f"{supplier_name} - match method", config["match_method"]),
        ]

        if config["match_method"] == "EAN + SKU":
            supplier_rows.extend(
                [
                    (f"{supplier_name} - EAN column", config["ean_column"]),
                    (f"{supplier_name} - SKU column", config["sku_column"]),
                    (f"{supplier_name} - priority", "EAN first, then SKU fallback"),
                ]
            )
        elif config["match_method"] == "EAN":
            supplier_rows.append((f"{supplier_name} - EAN column", config["ean_column"]))
        else:
            supplier_rows.append((f"{supplier_name} - SKU column", config["sku_column"]))

        supplier_rows.extend(
            [
                (f"{supplier_name} - price column", config["price_column"]),
                (f"{supplier_name} - matched products", int(full_result[price_col].notna().sum())),
                (
                    f"{supplier_name} - cheaper than our current price",
                    int(
                        (
                            full_result[price_col].notna()
                            & full_result["Our Price"].notna()
                            & (full_result[price_col] < full_result["Our Price"])
                        ).sum()
                    ),
                ),
                (f"{supplier_name} - duplicated identifiers", duplicate_info.get(supplier_name, 0)),
            ]
        )
        summary_rows.extend(supplier_rows)

    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        export_result.to_excel(writer, sheet_name="Price Comparison", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Price Comparison"]
        summary_ws = writer.sheets["Summary"]

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "bg_color": "#0F766E",
                "font_color": "#FFFFFF",
            }
        )
        money_fmt = workbook.add_format({"num_format": '€#,##0.00;[Red]-€#,##0.00'})
        percentage_fmt = workbook.add_format({"num_format": "0.00%"})
        decimal_fmt = workbook.add_format({"num_format": "#,##0.00"})
        integer_fmt = workbook.add_format({"num_format": "0"})
        supplier_cheaper_fmt = workbook.add_format({"bg_color": "#DCFCE7", "font_color": "#166534"})
        cheapest_fmt = workbook.add_format(
            {"bg_color": "#FEF3C7", "font_color": "#92400E", "bold": True}
        )
        not_found_fmt = workbook.add_format({"bg_color": "#F1F5F9", "font_color": "#475569"})

        for col_idx, name in enumerate(export_result.columns):
            worksheet.write(0, col_idx, name, header_fmt)

        widths = {
            "EAN": 17,
            "SKU": 22,
            "Our Price": 14,
            "Cheapest Price": 18,
            "Cheapest Supplier": 22,
            "Saving €": 13,
            "Saving %": 12,
            "Status": 19,
            "Matched Suppliers": 17,
            "Realisation Summ": 18,
            "Net Available Qty": 18,
            "Days Since Last Sale": 21,
            "MinStock": 12,
            "MaxStock": 12,
        }
        positions = {name: export_result.columns.get_loc(name) for name in export_result.columns}

        for col_idx, name in enumerate(export_result.columns):
            worksheet.set_column(col_idx, col_idx, widths.get(name, 18))

        for name in ["Our Price", *supplier_price_columns, "Cheapest Price", "Saving €"]:
            if name in positions:
                i = positions[name]
                worksheet.set_column(i, i, widths.get(name, 18), money_fmt)

        if "Saving %" in positions:
            i = positions["Saving %"]
            worksheet.set_column(i, i, widths["Saving %"], percentage_fmt)

        if "Realisation Summ" in positions:
            i = positions["Realisation Summ"]
            worksheet.set_column(i, i, widths["Realisation Summ"], decimal_fmt)

        for name in ["Matched Suppliers", "Net Available Qty", "Days Since Last Sale", "MinStock", "MaxStock"]:
            if name in positions:
                i = positions[name]
                worksheet.set_column(i, i, widths.get(name, 15), integer_fmt)

        worksheet.freeze_panes(1, 3)

        if len(export_result):
            worksheet.autofilter(0, 0, len(export_result), len(export_result.columns) - 1)
            our_price_letter = xl_col_to_name(positions["Our Price"])

            for supplier_col in supplier_price_columns:
                supplier_idx = positions[supplier_col]
                supplier_letter = xl_col_to_name(supplier_idx)
                worksheet.conditional_format(
                    1,
                    supplier_idx,
                    len(export_result),
                    supplier_idx,
                    {
                        "type": "formula",
                        "criteria": (
                            f'=AND({supplier_letter}2<>"",'
                            f'${our_price_letter}2<>"",'
                            f'{supplier_letter}2<${our_price_letter}2)'
                        ),
                        "format": supplier_cheaper_fmt,
                    },
                )

            for name in ["Cheapest Price", "Cheapest Supplier"]:
                idx = positions[name]
                worksheet.conditional_format(
                    1, idx, len(export_result), idx,
                    {"type": "no_blanks", "format": cheapest_fmt},
                )

            status_idx = positions["Status"]
            status_letter = xl_col_to_name(status_idx)
            worksheet.conditional_format(
                1,
                0,
                len(export_result),
                len(export_result.columns) - 1,
                {
                    "type": "formula",
                    "criteria": f'=${status_letter}2="NOT FOUND"',
                    "format": not_found_fmt,
                },
            )

        summary_ws.set_column("A:A", 50)
        summary_ws.set_column("B:B", 55)
        for col_idx, name in enumerate(summary_df.columns):
            summary_ws.write(0, col_idx, name, header_fmt)
        summary_ws.freeze_panes(1, 0)

    buffer.seek(0)
    return buffer.getvalue()


def style_browser_table(df, supplier_price_columns):
    target_blue = "background-color: #e0f2fe; color: #075985; font-weight: 800;"

    def style_row(row):
        styles = pd.Series("", index=row.index, dtype=object)
        target_price = row.get("Target Price", np.nan)
        has_target = pd.notna(target_price)

        if row.get("Lookup Status", "") == "CODE NOT FOUND":
            styles[:] = "background-color: #f1f5f9; color: #475569;"
            if "Target Price" in styles.index and has_target:
                styles["Target Price"] = target_blue
            return styles

        our_price = row.get("Our Price", np.nan)
        cheapest_price = row.get("Cheapest Price", np.nan)
        cheapest_supplier = row.get("Cheapest Supplier", "")

        if has_target:
            if "Target Price" in styles.index:
                styles["Target Price"] = target_blue

            for col in ["Our Price", *supplier_price_columns, "Cheapest Price"]:
                if col in styles.index:
                    value = row.get(col, np.nan)
                    if pd.notna(value) and float(value) < float(target_price):
                        styles[col] = target_blue

            below_target = row.get("Below Target", "")
            if "Below Target" in styles.index and pd.notna(below_target) and str(below_target).strip():
                styles["Below Target"] = target_blue
            return styles

        for col in supplier_price_columns:
            supplier_price = row.get(col, np.nan)
            if pd.notna(supplier_price) and pd.notna(our_price) and float(supplier_price) < float(our_price):
                styles[col] = "background-color: #dcfce7; color: #166534; font-weight: 600;"

        if pd.notna(cheapest_price):
            if "Cheapest Price" in styles.index:
                styles["Cheapest Price"] = "background-color: #fef3c7; color: #92400e; font-weight: 700;"
            if "Cheapest Supplier" in styles.index and cheapest_supplier:
                styles["Cheapest Supplier"] = "background-color: #fef3c7; color: #92400e; font-weight: 700;"

        return styles

    formatters = {
        "Our Price": "€{:.2f}",
        "Target Price": "€{:.2f}",
        "Cheapest Price": "€{:.2f}",
        "Saving €": "€{:.2f}",
        "Saving %": "{:.2%}",
        "Realisation Summ": "{:.2f}",
    }
    for col in supplier_price_columns:
        formatters[col] = "€{:.2f}"
    formatters = {k: v for k, v in formatters.items() if k in df.columns}

    return df.style.apply(style_row, axis=1).format(formatters, na_rep="")
