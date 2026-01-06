import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import tempfile

# ================= CONFIG =================

FONT_REGULAR = r"D:\papapdf\Noto_Sans_Devanagari\static\NotoSansDevanagari-Regular.ttf"
FONT_BOLD = r"D:\papapdf\Noto_Sans_Devanagari\static\NotoSansDevanagari-SemiBold.ttf"

TITLE_FONT_SIZE = 14
COLUMN_FONT_SIZE = 11
DATA_FONT_SIZE = 10

ROW_HEIGHT = 25
SCALE = 2

LEFT_MARGIN = 20
BOTTOM_MARGIN = 10

PAGE_WIDTH, PAGE_HEIGHT = A4

HEADERS = [
    "बूथ नं.", "अ क्र.", "आडनाव", "नाव",
    "वडिलांचे/पतीचे नाव", "मतदान कार्ड",
    "वय", "लिंग", "विधानसभा क्र"
]

COL_INDEX = [3, 4, 6, 7, 8, 17, 15, 14, 18]
COL_WIDTHS = [38, 33, 63, 90, 100, 85, 35, 35, 70]
TABLE_WIDTH = sum(COL_WIDTHS)

# =========================================

def draw_table_row(texts, font, bold=False):
    img = Image.new(
        "RGB",
        (TABLE_WIDTH * SCALE, ROW_HEIGHT * SCALE),
        "white"
    )
    draw = ImageDraw.Draw(img)

    x = 0
    for i, txt in enumerate(texts):
        x_scaled = x * SCALE
        w_scaled = COL_WIDTHS[i] * SCALE

        draw.rectangle(
            [x_scaled, 0, x_scaled + w_scaled, ROW_HEIGHT * SCALE],
            outline="black",
            width=2 * SCALE if bold else 1 * SCALE
        )

        text_y = (ROW_HEIGHT * SCALE - font.size) // 2
        draw.text(
            (x_scaled + 6 * SCALE, text_y),
            str(txt),
            fill="black",
            font=font
        )

        x += COL_WIDTHS[i]

    return img


def generate_booth_pdfs(main_excel, output_dir, ward_no):
    df = pd.read_excel(main_excel, dtype=str).fillna("")
    df = df.sort_values(by=[df.columns[9], df.columns[10]], kind="mergesort")

    booth_groups = df.groupby(df.columns[3])
    os.makedirs(output_dir, exist_ok=True)

    title_font = ImageFont.truetype(FONT_REGULAR, TITLE_FONT_SIZE * SCALE)
    header_font = ImageFont.truetype(FONT_BOLD, COLUMN_FONT_SIZE * SCALE)
    data_font = ImageFont.truetype(FONT_REGULAR, DATA_FONT_SIZE * SCALE)

    for booth_no, booth_df in booth_groups:
        pdf_path = os.path.join(output_dir, f"Booth_{booth_no}.pdf")
        c = canvas.Canvas(pdf_path, pagesize=A4)

        y = PAGE_HEIGHT - 90
        row_counter = 0
        page_counter = 1

        with tempfile.TemporaryDirectory() as tmp:

            def draw_column_header():
                nonlocal y, page_counter
                img = draw_table_row(HEADERS, header_font, bold=True)
                path = os.path.join(tmp, f"header_{page_counter}.png")
                img.save(path)
                c.drawImage(path, LEFT_MARGIN, y, width=TABLE_WIDTH, height=ROW_HEIGHT)
                y -= ROW_HEIGHT

            def new_page():
                nonlocal y, page_counter
                c.showPage()
                page_counter += 1
                y = PAGE_HEIGHT - 40
                draw_column_header()

            # ---------- MAIN HEADER (FIRST PAGE ONLY) ----------
            title_img = Image.new(
                "RGB",
                (int(PAGE_WIDTH * SCALE), int(80 * SCALE)),
                "white"
            )
            d = ImageDraw.Draw(title_img)

            t1 = "महाराष्ट्र महानगरपालिका निवडणूक २०२५"
            t2 = f"जळगाव महानगरपालिका ({ward_no}) प्रभाग {ward_no}"

            x1 = (PAGE_WIDTH * SCALE - d.textlength(t1, title_font)) // 2
            x2 = (PAGE_WIDTH * SCALE - d.textlength(t2, title_font)) // 2

            d.text((x1, 5 * SCALE), t1, font=title_font, fill="black")
            d.text((x2, 40 * SCALE), t2, font=title_font, fill="black")

            title_path = os.path.join(tmp, "title.png")
            title_img.save(title_path)

            c.drawImage(title_path, 0, PAGE_HEIGHT - 80, width=PAGE_WIDTH, height=80)

            draw_column_header()

            # ---------- DATA ROWS ----------
            for _, row in booth_df.iterrows():

                if y < BOTTOM_MARGIN:
                    new_page()

                values = [row.iloc[i] for i in COL_INDEX]
                img = draw_table_row(values, data_font)

                row_counter += 1
                path = os.path.join(tmp, f"row_{row_counter}.png")
                img.save(path)

                c.drawImage(path, LEFT_MARGIN, y, width=TABLE_WIDTH, height=ROW_HEIGHT)
                y -= ROW_HEIGHT

        c.save()
        print(f"✅ Created PDF: {pdf_path}")


# ================= RUN =================
if __name__ == "__main__":
    generate_booth_pdfs(
        main_excel=r"D:\papa\5.1.26\sample excel.xlsx",
        output_dir=r"D:\papa\5.1.26",
        ward_no="17"
    )
