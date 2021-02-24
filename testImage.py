from PIL import Image
import xlsxwriter
import os
import sys

# Print iterations progress
def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total: 
        print()


def rgb2hex(rgb: tuple):
    return '#{:02x}{:02x}{:02x}'.format(rgb[0], rgb[1], rgb[2])

def processImage(imageName):
    fileName = os.path.splitext(imageName)[0]
    im = Image.open( f'./images/{imageName}') # Can be many different formats.
    pix = im.load() # Loading image
    picSize = im.size  # Get the width and hight of the image for iterating over

    # creating workbook
    file_path = f'./arts/{fileName}.xlsx'  # path to save workbook
    workbook = xlsxwriter.Workbook(file_path)  # Initializing workbook
    art_sheet = workbook.add_worksheet(fileName)  # worksheet sheet name
    art_sheet.set_column(0, picSize[0], 1)  # setting artsheet column width
    colorMap = {}  # workbook format map
    #writing colors in excel
    printProgressBar(0, picSize[1], prefix = 'Drawing:', suffix = 'Complete', length = 50)
    for y in range(0, picSize[1]):
        printProgressBar(y + 1, picSize[1], prefix = 'Drawing:', suffix = 'Complete', length = 50)
        art_sheet.set_row(y, 10)  # setting artsheet row height
        for x in range(0, picSize[0]):
            hexColor = rgb2hex(pix[x,y])  # Get the RGBA Value of the a pixel of an image
            colorMap[hexColor] = workbook.add_format({'bg_color':hexColor})  #setting color map
            art_sheet.write(y, x,'',colorMap[hexColor])
    workbook.close()
    print("\nArt Sheet is saved Successfully...\n\n")

if __name__ == "__main__":
    if len(sys.argv) == 2:
        processImage(str(sys.argv[1]))
    else:
        print("Provide a file first")