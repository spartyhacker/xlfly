import xlwings as xw
import os


def open_bk(src: str):
    "shortcut to have src workbook object"
    return xw.books.open(src)


def rng_to_link(rng: xw.Range) -> str:

    # due to lack of parent object in xlwings
    # I had to use api to get original object names
    sht_name = rng.api.Parent.Name
    wb_name = rng.api.Parent.Parent.Name
    wb = xw.books[wb_name]
    sht = wb.sheets[sht_name]
    origin = sht.range((rng.row, rng.column))

    nrow, ncol = rng.shape

    ranges = []

    for row in range(0, nrow):
        ranges_unit = []
        for col in range(0, ncol):
            # value link
            # 'D:\coding\[src_data.xlsx]Sheet1'!A1

            unit_cell = origin.offset(row, col)
            full_address = unit_cell.get_address(
                include_sheetname=True, column_absolute=False, row_absolute=False
            )

            sheet_name, cell_name = full_address.split("!")
            wb_path = os.path.dirname(wb.fullname)
            link_path = f"'{wb_path}\\[{wb.name}]{sheet_name}'!{cell_name}"

            ranges_unit.append("=" + link_path)

        ranges.append(ranges_unit)

    return ranges


def to_link(self: xw.Range):
    return rng_to_link(self)


setattr(xw.Range, "to_link", to_link)


def main():
    src = r"C:\Users\Tony\OneDrive\Documents\src_data.xlsx"
    wb = xw.books.open(src)
    address = wb.sheets["Sheet1"]["A1"].expand("table").get_address()
    pass


if __name__ == "__main__":
    main()
