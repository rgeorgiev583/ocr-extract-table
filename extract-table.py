#!/usr/bin/env python
from PIL import Image, ImageOps
import subprocess
import sys
import os
import glob
import re
import tempfile
import argparse

# minimum run of adjacent pixels to call something a line
HORIZ_PIXEL_COUNT_THRESHOLD = 300
VERT_PIXEL_COUNT_THRESHOLD = 300

args = None


def get_horiz_line_coords(pixmap, width, height):
    """Get start/end coordinates of lines containing a horizontal run of at least HORIZ_PIXEL_COUNT_THRESHOLD pixels of the border color"""
    horiz_line_coords = []
    for y in range(height):
        x1, x2 = (None, None)
        border_color_pixel_count = 0
        max_border_color_pixel_count = 0
        for x in range(width):
            if pixmap[x, y] == args.border_color:
                border_color_pixel_count = border_color_pixel_count + 1
                if not x1:
                    x1 = x
                x2 = x
            else:
                if border_color_pixel_count > max_border_color_pixel_count:
                    max_border_color_pixel_count = border_color_pixel_count
                border_color_pixel_count = 0
        if max_border_color_pixel_count > HORIZ_PIXEL_COUNT_THRESHOLD:
            horiz_line_coords.append((x1, y, x2, y))
    return horiz_line_coords


def get_vert_line_coords(pixmap, width, height):
    """Get start/end coordinates of lines containing a vertical run of at least VERT_PIXEL_COUNT_THRESHOLD pixels of the border color"""
    vert_line_coords = []
    for x in range(width):
        y1, y2 = (None, None)
        border_color_pixel_count = 0
        max_border_color_pixel_count = 0
        for y in range(height):
            if pixmap[x, y] == args.border_color:
                border_color_pixel_count = border_color_pixel_count + 1
                if not y1:
                    y1 = y
                y2 = y
            else:
                if border_color_pixel_count > max_border_color_pixel_count:
                    max_border_color_pixel_count = border_color_pixel_count
                border_color_pixel_count = 0
        if max_border_color_pixel_count > VERT_PIXEL_COUNT_THRESHOLD:
            vert_line_coords.append((x, y1, x, y2))
    return vert_line_coords


def get_col_coords(vert_line_coords):
    """Get top-left and bottom-right coordinates for each column from a list of vertical lines"""
    col_coords = []
    for i in range(1, len(vert_line_coords)):
        if vert_line_coords[i][0] - vert_line_coords[i-1][0] > 1:
            col_coords.append((vert_line_coords[i-1][0], vert_line_coords[i-1][1],
                               vert_line_coords[i][2], vert_line_coords[i][3]))
    return col_coords


def get_row_coords(horiz_line_coords):
    """Get top-left and bottom-right coordinates for each row from a list of horizontal lines"""
    row_coords = []
    for i in range(1, len(horiz_line_coords)):
        if horiz_line_coords[i][1] - horiz_line_coords[i-1][3] > 1:
            row_coords.append((horiz_line_coords[i-1][0], horiz_line_coords[i-1][1],
                               horiz_line_coords[i][2], horiz_line_coords[i][3]))
    return row_coords


def get_cell_coords(row_coords, col_coords):
    """Get top-left and bottom-right coordinates for each cell using row and column coordinates"""
    cell_coords = {}
    for i, row_coord in enumerate(row_coords):
        cell_coords.setdefault(i, {})
        for j, col_coord in enumerate(col_coords):
            x1 = col_coord[0]
            y1 = row_coord[1]
            x2 = col_coord[2]
            y2 = row_coord[3]
            cell_coords[i][j] = (x1, y1, x2, y2)
    return cell_coords


def ocr_table_cell(image, cell_coords, row_num, col_num):
    """Return OCRed text from the specified table cell"""
    base_path = os.path.join(args.workdir, "{}-{}".format(row_num, col_num))
    image_file = "{}.tif".format(base_path)
    text_file = "{}.txt".format(base_path)
    ocr_cmd = ["tesseract", "-l", args.language, image_file, base_path]
    # extract cell from whole image, grayscale (1-color channel), monochrome
    cell_region = image.crop(cell_coords[row_num][col_num])
    cell_region = ImageOps.grayscale(cell_region)
    cell_region = cell_region.point(lambda p: p > 200 and 255)
    # determine background color (most used color)
    cell_histogram = cell_region.histogram()
    if cell_histogram[0] > cell_histogram[255]:
        bg_color = 0
    else:
        bg_color = 255
    # trim borders by finding top-left and bottom-right bg pixels
    cell_pixmap = cell_region.load()
    x1, y1 = 0, 0
    x2, y2 = cell_region.size
    x2, y2 = x2-1, y2-1
    while cell_pixmap[x1, y1] != bg_color:
        x1 += 1
        y1 += 1
    while cell_pixmap[x2, y2] != bg_color:
        x2 -= 1
        y2 -= 1
    # save as TIFF and extract text with Tesseract OCR
    cell_region = cell_region.crop((x1, y1, x2, y2))
    cell_region.save(image_file, "TIFF")
    subprocess.call(ocr_cmd, stderr=subprocess.PIPE)
    lines = [line.strip() for line in open(text_file).readlines()]
    return "\n".join(filter(lambda line: line != "", lines))


def get_image_table_data(filename):
    """Extract textual table data from a spreadsheet-like image file"""
    image = Image.open(filename)
    pixmap = image.load()
    width, height = image.size
    horiz_line_coords = get_horiz_line_coords(pixmap, width, height)
    sys.stderr.write("{}: horizontal line count: {}\n".format
                     (filename, len(horiz_line_coords)))
    vert_line_coords = get_vert_line_coords(pixmap, width, height)
    sys.stderr.write("{}: vertical line count: {}\n".format
                     (filename, len(vert_line_coords)))
    row_coords = get_row_coords(horiz_line_coords)
    sys.stderr.write("{}: row count: {}\n".format(filename, len(row_coords)))
    col_coords = get_col_coords(vert_line_coords)
    sys.stderr.write("{}: column count: {}\n".format(
        filename, len(col_coords)))
    cell_coords = get_cell_coords(row_coords, col_coords)
    table_data = []
    for row_num in range(len(row_coords)):
        table_data.append([ocr_table_cell(image, cell_coords, row_num, col_num)
                      for col_num in range(len(col_coords))])
    return table_data


def split_pdf_into_pngs(pdf_file_path):
    """Split each page of the PDF file into a PNG file and return their paths"""
    pdf_file_path_stem, _ = os.path.splitext(os.path.basename(pdf_file_path))
    png_base_path = os.path.join(args.workdir, pdf_file_path_stem)
    subprocess.call(["pdftoppm", pdf_file_path, png_base_path, "-png"])
    return [png_file_path for png_file_path in glob.glob(os.path.join(args.workdir, "{}-*".format(png_base_path)))]


def extract_pdf_table(pdf_file_path):
    """Extract table data from a PDF file"""
    png_file_paths = split_pdf_into_pngs(pdf_file_path)
    sys.stderr.write("Page count: {}\n".format(len(png_file_paths)))
    # extract table data from each page
    table = []
    for png_file_path in png_file_paths:
        table_data = get_image_table_data(png_file_path)
        for row in table_data:
            table.append(row)
    return table


def extract_pdf_table_into_str(pdf_file_path):
    """Extract table data from a PDF file and convert it to a string"""
    table = extract_pdf_table(pdf_file_path)
    lines = [args.csv_separator.join(row) for row in table]
    return "\n".join(lines)


def color(string):
    value_match = re.fullmatch(
        r"\s*#([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})\s*", string)
    if value_match is None:
        value_match = re.fullmatch(
            r"\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)\s*", string)
    if value_match is None:
        errmsg = "{} does not represent a valid color".format(string)
        raise argparse.ArgumentTypeError(errmsg)
    value = tuple(map(lambda channel_value: int(
        channel_value), value_match.groups()))
    return value


if __name__ == "__main__":
    arg_parser = argparse.ArgumentParser(
        description="Extract table data using OCR from a PDF into CSV format.")
    arg_parser.add_argument("input_files", metavar="FILE", nargs="+",
                            help="PDF file to extract from")
    arg_parser.add_argument("-l", "--language", default="eng",
                            help="the language of the data (default: 'eng')")
    arg_parser.add_argument("-b", "--border-color", type=color, default=(0, 0, 0),
                            help="the (main) color of the borders of the table. Can be expressed as a hexadecimal or decimal RGB triple. For instance, the color black is represented as #000000 in hex and (0, 0, 0) in decimal (default: '#000000')")
    arg_parser.add_argument("-s", "--csv-separator", default=",",
                            help="the separator character used for the CSV output (default: ',')")
    arg_parser.add_argument(
        "-w", "--workdir", help="path to the working directory where intermediate files are placed (default: a temporary directory)")
    args = arg_parser.parse_args()

    for input_file in args.input_files:
        if args.workdir is None:
            with tempfile.TemporaryDirectory() as tempdir:
                args.workdir = tempdir
                print(extract_pdf_table_into_str(input_file))
        else:
            print(extract_pdf_table_into_str(input_file))
