"""Build the ISSB (IFRS S1 / IFRS S2) sustainability-reporting working-paper workbook.

Usage:  python3 build_workpapers.py [output.xlsx]
Then recalculate once (LibreOffice / Excel) so formula results are cached.
"""
import sys
from wp_lib import Book
from wp_part1 import build_part1
from wp_part2 import build_part2
from wp_part3 import build_part3

OUT = sys.argv[1] if len(sys.argv) > 1 else 'TTM_ISSB_Sustainability_Working_Papers_FY2025.xlsx'


def main():
    X = Book(OUT)
    X.wb.set_properties({'title': 'ISSB sustainability reporting working papers (IFRS S1 / IFRS S2)',
                         'subject': 'Illustrative sample client - garment manufacturer, Cambodia',
                         'comments': 'Fictional data for teaching purposes'})
    X.dash = X.wb.add_worksheet('Dashboard')
    X.dash.set_tab_color('#0B2545')
    X.idx = X.wb.add_worksheet('Index')
    X.idx.set_tab_color('#0B2545')
    build_part1(X)
    build_part2(X)
    build_part3(X)
    X.wb.close()
    print('written', OUT, '| sheets:', len(X.wb.worksheets()), '| names:', len(X.names))


if __name__ == '__main__':
    main()
