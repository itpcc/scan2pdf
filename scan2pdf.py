#!/usr/bin/env python

"""scan2pdf.py: Little script to help scanning document and generate a PDF using WIA-compatible scanner and Pillow"""
__license__ = "MIT"

import win32com.client
from PIL import Image
from pathlib import Path

WIA_IMG_FORMAT_BMP = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
WIA_COMMAND_TAKE_PICTURE = "{9B26B7B2-ACAD-11D2-A093-00C04F72DC3C}"

WIA_CONST_SCAN_COLOR_BITDEPTH = "4104"
WIA_CONST_SCAN_COLOR_MODE = "6146"
WIA_CONST_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147"
WIA_CONST_VERTICAL_SCAN_RESOLUTION_DPI = "6148"
WIA_CONST_HORIZONTAL_SCAN_START_PIXEL = "6149"
WIA_CONST_VERTICAL_SCAN_START_PIXEL = "6150"
WIA_CONST_HORIZONTAL_SCAN_SIZE_PIXELS = "6151"
WIA_CONST_VERTICAL_SCAN_SIZE_PIXELS = "6152"
WIA_CONST_SCAN_BRIGHTNESS_PERCENTS = "6154"
WIA_CONST_SCAN_CONTRAST_PERCENTS = "6155"

def scanToImage (filePrefix, pageNo):
    # Call WIA dialog
    # @see https://sites.tntech.edu/renfro/2009/09/03/capturing-an-image-from-a-wia-compatible-digital-camera/
    wia = win32com.client.Dispatch("WIA.CommonDialog")
    dev = wia.ShowSelectDevice()
    for command in dev.Commands:
        if command.CommandID == WIA_COMMAND_TAKE_PICTURE:
            print("Page {}: Scanning".format(pageNo))
            dev.ExecuteCommand(WIA_COMMAND_TAKE_PICTURE)

    # Change scanner properties
    # @see https://stackoverflow.com/a/55050563
    # @see https://csharp.hotexamples.com/examples/-/CommonDialog/ShowTransfer/php-commondialog-showtransfer-method-examples.html
    scanner = dev.Items[dev.Items.Count]
    # Color 24 bits depth...
    scanner.Properties[WIA_CONST_SCAN_COLOR_BITDEPTH] = 24
    # ... in color
    scanner.Properties[WIA_CONST_SCAN_COLOR_MODE] = 1
    # Scan at 300 DPI
    scanner.Properties[WIA_CONST_HORIZONTAL_SCAN_RESOLUTION_DPI] = 300
    scanner.Properties[WIA_CONST_VERTICAL_SCAN_RESOLUTION_DPI] = 300
    # Set scanning offset
    scanner.Properties[WIA_CONST_HORIZONTAL_SCAN_START_PIXEL] = 0
    scanner.Properties[WIA_CONST_VERTICAL_SCAN_START_PIXEL] = 0
    # Max: 2550
    scanner.Properties[WIA_CONST_HORIZONTAL_SCAN_SIZE_PIXELS] = 2500
    # Max: 3500
    scanner.Properties[WIA_CONST_VERTICAL_SCAN_SIZE_PIXELS] = 3500
    # Possible values: -20 to 20
    scanner.Properties[WIA_CONST_SCAN_BRIGHTNESS_PERCENTS] = 5
    scanner.Properties[WIA_CONST_SCAN_CONTRAST_PERCENTS] = -20

    # Do the actual scanning
    print("Page {}: Transfering".format(pageNo))
    image = scanner.Transfer(WIA_IMG_FORMAT_BMP)

    # Due to limitation of the interface, only BMP file is available. (You can convert using Pillow though)
    print("Page {}: Saving".format(pageNo))
    filePath = "./images/img-{}-{}.bmp".format(filePrefix, pageNo)
    image.SaveFile(filePath)
    return filePath

filePrefix = input("Enter subject code: ")
pageNo = int(input("Enter page number: "))
fileList = []

Path("./images").mkdir(parents=True, exist_ok=True)

for i in range(1, pageNo + 1):
    input("Press enter to scan page {} ...".format(i))
    filePath = scanToImage(filePrefix, i)
    fileList.append(Image.open(filePath))

firstPage = fileList.pop()
# Save PDF file
# @see https://stackoverflow.com/a/63436357
# @see https://pillow.readthedocs.io/en/latest/handbook/image-file-formats.html#pdf
firstPage.save(
    "./LAW{}-<STD Code>.pdf".format(filePrefix),
    "PDF" , resolution=300, save_all=True, append_images=fileList,
    title="สมุดคำตอบวิชา LAW{} <YourName>".format(filePrefix)
)
